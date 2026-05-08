"""COM wrapper around Outlook.Application.

All public functions return plain-Python dicts/lists suitable for JSON serialization.
Raises `OutlookError` for user-facing failures.
"""
from __future__ import annotations

import ctypes
import datetime as _dt
import re
from typing import Any, Iterable, Literal, Sequence

import pythoncom
import pywintypes
import win32com.client

# Outlook enum constants (avoid importing win32com.client.constants at top level —
# it only populates after an EnsureDispatch call).
OL_FOLDER_DELETED = 3
OL_FOLDER_OUTBOX = 4
OL_FOLDER_SENT = 5
OL_FOLDER_INBOX = 6
OL_FOLDER_CALENDAR = 9
OL_FOLDER_CONTACTS = 10
OL_FOLDER_TASKS = 13
OL_FOLDER_DRAFTS = 16
OL_FOLDER_JUNK = 23

OL_MAIL_ITEM = 0
OL_APPT_ITEM = 1

OL_MEETING_RESPONSE = {
    "accept": 3,      # olMeetingAccepted
    "tentative": 2,   # olMeetingTentative
    "decline": 4,     # olMeetingDeclined
}

OL_IMPORTANCE = {"low": 0, "normal": 1, "high": 2}
OL_IMPORTANCE_INV = {v: k for k, v in OL_IMPORTANCE.items()}

_NAMED_FOLDERS = {
    "inbox": OL_FOLDER_INBOX,
    "sent": OL_FOLDER_SENT,
    "sent items": OL_FOLDER_SENT,
    "drafts": OL_FOLDER_DRAFTS,
    "deleted": OL_FOLDER_DELETED,
    "deleted items": OL_FOLDER_DELETED,
    "trash": OL_FOLDER_DELETED,
    "junk": OL_FOLDER_JUNK,
    "junk email": OL_FOLDER_JUNK,
    "outbox": OL_FOLDER_OUTBOX,
    "calendar": OL_FOLDER_CALENDAR,
    "contacts": OL_FOLDER_CONTACTS,
    "tasks": OL_FOLDER_TASKS,
    # German (de) aliases
    "posteingang": OL_FOLDER_INBOX,
    "gesendete elemente": OL_FOLDER_SENT,
    "gesendet": OL_FOLDER_SENT,
    "entwürfe": OL_FOLDER_DRAFTS,
    "gelöschte elemente": OL_FOLDER_DELETED,
    "papierkorb": OL_FOLDER_DELETED,
    "junk-e-mail": OL_FOLDER_JUNK,
    "postausgang": OL_FOLDER_OUTBOX,
    "kalender": OL_FOLDER_CALENDAR,
    "kontakte": OL_FOLDER_CONTACTS,
    "aufgaben": OL_FOLDER_TASKS,
}


class OutlookError(Exception):
    """User-facing error from the Outlook wrapper."""


_app = None
_ns = None


def _app_ns():
    global _app, _ns
    if _app is None:
        pythoncom.CoInitialize()
        _app = win32com.client.Dispatch("Outlook.Application")
        _ns = _app.GetNamespace("MAPI")
    return _app, _ns


def _to_iso(val: Any) -> str | None:
    if val is None:
        return None
    if isinstance(val, pywintypes.TimeType):
        # pywintypes.datetime is tz-aware; normalize to ISO
        return _dt.datetime(
            val.year, val.month, val.day, val.hour, val.minute, val.second,
            tzinfo=val.tzinfo,
        ).isoformat()
    if isinstance(val, _dt.datetime):
        return val.isoformat()
    return str(val)


def _parse_dt(val: str | _dt.datetime) -> _dt.datetime:
    if isinstance(val, _dt.datetime):
        return val
    try:
        return _dt.datetime.fromisoformat(val)
    except ValueError as e:
        raise OutlookError(f"Invalid datetime (expected ISO 8601): {val!r}") from e


def _format_filter_dt(dt: _dt.datetime) -> str:
    """Format a datetime for Outlook DASL/Jet filter strings.

    Outlook COM filters expect US-style dates: ``M/D/YYYY h:mm AM/PM``.
    Using ``strftime("%p")`` is broken on non-English Windows locales
    (e.g. German returns an empty string instead of AM/PM), so we build
    the 12-hour period string manually.
    """
    h = dt.hour
    period = "AM" if h < 12 else "PM"
    h12 = h % 12 or 12
    return f"{dt.month}/{dt.day}/{dt.year} {h12}:{dt.minute:02d} {period}"


def _get_windows_short_date_format() -> str:
    """Return the Windows short date format string (e.g. 'dd.MM.yyyy' on German).

    Falls back to US-style ``M/d/yyyy`` on non-Windows platforms.
    """
    try:
        LOCALE_USER_DEFAULT = 0x0400
        LOCALE_SSHORTDATE = 0x1F
        buf = ctypes.create_unicode_buffer(80)
        ctypes.windll.kernel32.GetLocaleInfoW(  # type: ignore[attr-defined]
            LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, buf, 80,
        )
        return buf.value or "M/d/yyyy"
    except (AttributeError, OSError):
        # Not on Windows (no ctypes.windll) — default to US format.
        return "M/d/yyyy"


# Cache the format at module load so we don't call into kernel32 on every query.
_WIN_SHORT_DATE_FMT: str = _get_windows_short_date_format()


def _format_jet_filter_dt(dt: _dt.datetime) -> str:
    """Format a datetime for Outlook **Jet** filter strings (calendar).

    Jet filters (``[Start] >= '…'``) require the date in the **Windows
    system locale** format — e.g. ``dd.MM.yyyy`` on German Windows,
    ``M/d/yyyy`` on US Windows. Using the wrong format silently returns
    zero results instead of raising an error.

    DASL filters (``@SQL="urn:…" >= '…'``) accept US-style dates and
    should use :func:`_format_filter_dt` instead.
    """
    fmt = _WIN_SHORT_DATE_FMT
    # Windows date-format tokens → Python values.
    # Replace longer tokens first to avoid partial matches (dd before d).
    result = fmt
    result = result.replace("dddd", "")   # skip day-of-week name
    result = result.replace("ddd", "")    # skip abbreviated day name
    result = result.replace("dd", f"{dt.day:02d}")
    result = result.replace("d", str(dt.day))
    result = result.replace("MMMM", "")   # skip month name
    result = result.replace("MMM", "")    # skip abbreviated month name
    result = result.replace("MM", f"{dt.month:02d}")
    result = result.replace("M", str(dt.month))
    result = result.replace("yyyy", str(dt.year))
    result = result.replace("yy", str(dt.year % 100).zfill(2))
    return f"{result} {dt.hour:02d}:{dt.minute:02d}"


def _escape_dasl(val: str) -> str:
    """Escape a string for use inside a DASL @SQL= LIKE filter.

    - Single quotes are doubled (standard SQL escaping).
    - LIKE wildcards (% and _) are escaped so user input is treated as
      literal text and cannot alter the filter structure.
    """
    return val.replace("'", "''").replace("%", "[%]").replace("_", "[_]")


_BLOCKED_ATTACHMENT_PATTERNS: list[re.Pattern] = [
    re.compile(r"[/\\]\.ssh[/\\]", re.IGNORECASE),
    re.compile(r"[/\\]\.gnupg[/\\]", re.IGNORECASE),
    re.compile(r"[/\\]\.aws[/\\]", re.IGNORECASE),
    re.compile(r"[/\\]\.azure[/\\]", re.IGNORECASE),
    re.compile(r"[/\\]\.kube[/\\]", re.IGNORECASE),
    re.compile(r"\.pem$", re.IGNORECASE),
    re.compile(r"\.key$", re.IGNORECASE),
    re.compile(r"\.pfx$", re.IGNORECASE),
    re.compile(r"\.p12$", re.IGNORECASE),
    re.compile(r"id_rsa", re.IGNORECASE),
    re.compile(r"id_ed25519", re.IGNORECASE),
    re.compile(r"credentials\.json$", re.IGNORECASE),
    re.compile(r"\.env$", re.IGNORECASE),
    re.compile(r"\.env\.", re.IGNORECASE),
    re.compile(r"\.netrc$", re.IGNORECASE),
]


def _validate_attachment_path(path: str) -> None:
    """Raise OutlookError if path looks like a sensitive credential file."""
    normalized = path.replace("\\", "/")
    for pattern in _BLOCKED_ATTACHMENT_PATTERNS:
        if pattern.search(normalized):
            raise OutlookError(
                f"Attachment blocked — path matches a sensitive file pattern: {path!r}"
            )


_DANGEROUS_HTML_RE = re.compile(
    r"<\s*(?:script|iframe|object|embed|applet|form|meta\s+http-equiv)[^>]*>",
    re.IGNORECASE,
)


def _sanitize_html_body(html: str) -> str:
    """Strip dangerous HTML tags from email body content.

    Removes <script>, <iframe>, <object>, <embed>, <applet>, <form>,
    and <meta http-equiv=...> tags that could be injected via prompt
    injection attacks through malicious email content.
    """
    return _DANGEROUS_HTML_RE.sub("<!-- blocked -->", html)


def _parse_categories(raw: str | None) -> list[str]:
    if not raw:
        return []
    # Outlook stores categories as a locale-delimited string. Comma is the most
    # common separator in en-US; semicolon shows up in some locales. Split on both.
    parts: list[str] = []
    for chunk in raw.split(","):
        parts.extend(chunk.split(";"))
    return [p.strip() for p in parts if p.strip()]


def _walk_subfolders(parent, sub_parts: list[str]):
    """Walk into subfolders by name, with case-insensitive fallback.

    Folders.Item() uses the exact localized name and raises com_error on
    mismatch. The fallback scans all children case-insensitively so that
    e.g. "Inbox" resolves on a German install where the folder is called
    "Posteingang" at the COM level.
    """
    folder = parent
    for sub in sub_parts:
        try:
            folder = folder.Folders.Item(sub)
        except pywintypes.com_error:
            sub_lower = sub.lower()
            matched = None
            for j in range(1, folder.Folders.Count + 1):
                child = folder.Folders.Item(j)
                if child.Name.lower() == sub_lower:
                    matched = child
                    break
            if matched is None:
                raise OutlookError(
                    f"Subfolder {sub!r} not found under {getattr(folder, 'Name', '?')!r}"
                )
            folder = matched
    return folder


def _resolve_folder(path: str | None):
    """Resolve a folder path to a COM Folder object.

    Accepted forms:
      - None or "" → Inbox
      - "inbox", "sent", "drafts", "deleted", "junk", "outbox", "calendar"
        (plus German equivalents: "posteingang", "entwürfe", etc.)
      - "inbox/Processed/Q1" → subfolder under default inbox
      - "account@example.com/Inbox/Processed" → walk from named store
    """
    _, ns = _app_ns()
    if not path:
        return ns.GetDefaultFolder(OL_FOLDER_INBOX)

    parts = [p for p in path.split("/") if p]
    if not parts:
        return ns.GetDefaultFolder(OL_FOLDER_INBOX)

    first_lower = parts[0].lower()

    # Single named folder shortcut
    if len(parts) == 1 and first_lower in _NAMED_FOLDERS:
        return ns.GetDefaultFolder(_NAMED_FOLDERS[first_lower])

    # Named folder prefix + subpath (e.g., "inbox/Processed")
    if first_lower in _NAMED_FOLDERS:
        folder = ns.GetDefaultFolder(_NAMED_FOLDERS[first_lower])
        return _walk_subfolders(folder, parts[1:])

    # Maybe first part is a store (account) name
    stores = ns.Folders
    for i in range(1, stores.Count + 1):
        store = stores.Item(i)
        if store.Name.lower() == first_lower:
            return _walk_subfolders(store, parts[1:])

    # Fall back: treat whole path as subpath under default store root
    folder = ns.DefaultStore.GetRootFolder()
    return _walk_subfolders(folder, parts)


def _get_item(entry_id: str, store_id: str | None = None):
    _, ns = _app_ns()
    try:
        if store_id:
            return ns.GetItemFromID(entry_id, store_id)
        return ns.GetItemFromID(entry_id)
    except pywintypes.com_error as e:
        raise OutlookError(f"Item not found (entry_id={entry_id!r}): {e}") from e


# PR_SMTP_ADDRESS — the canonical SMTP address on an AddressEntry
_PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"


def _address_entry_smtp(address_entry) -> str | None:
    """Resolve an AddressEntry to its SMTP address, handling Exchange users."""
    if address_entry is None:
        return None
    try:
        exu = address_entry.GetExchangeUser()
        if exu is not None:
            return exu.PrimarySmtpAddress
    except Exception:
        pass
    try:
        exdl = address_entry.GetExchangeDistributionList()
        if exdl is not None:
            return exdl.PrimarySmtpAddress
    except Exception:
        pass
    try:
        return address_entry.PropertyAccessor.GetProperty(_PR_SMTP_ADDRESS)
    except Exception:
        pass
    try:
        return address_entry.Address
    except Exception:
        return None


def _recipient_smtp(recipient) -> str | None:
    try:
        addr = recipient.Address
    except Exception:
        addr = None
    # Exchange legacy DN — try to resolve via AddressEntry
    if addr and (addr.startswith("/") or "/cn=" in addr.lower()):
        try:
            resolved = _address_entry_smtp(recipient.AddressEntry)
            if resolved and "@" in resolved:
                return resolved
        except Exception:
            pass
    return addr


def _recipients_to_list(recipients) -> list[dict]:
    out = []
    for i in range(1, recipients.Count + 1):
        r = recipients.Item(i)
        entry = {"name": r.Name, "email": _recipient_smtp(r), "type": None}
        try:
            entry["type"] = {1: "to", 2: "cc", 3: "bcc"}.get(r.Type)
        except Exception:
            pass
        out.append(entry)
    return out


def _sender_smtp(m) -> str | None:
    """Get the SMTP sender address, resolving Exchange legacy DNs."""
    sender_type = getattr(m, "SenderEmailType", None)
    addr = getattr(m, "SenderEmailAddress", None)
    if sender_type == "EX" or (addr and (addr.startswith("/") or "/cn=" in addr.lower())):
        try:
            return _address_entry_smtp(m.Sender)
        except Exception:
            pass
    return addr


def _email_summary(m) -> dict:
    try:
        preview = (m.Body or "")[:200].replace("\r\n", "\n").strip()
    except Exception:
        preview = ""
    return {
        "entry_id": m.EntryID,
        "store_id": getattr(m, "StoreID", None),
        "subject": m.Subject,
        "sender_name": getattr(m, "SenderName", None),
        "sender_email": _sender_smtp(m),
        "received_time": _to_iso(getattr(m, "ReceivedTime", None)),
        "sent_time": _to_iso(getattr(m, "SentOn", None)),
        "unread": bool(getattr(m, "UnRead", False)),
        "has_attachments": getattr(m, "Attachments", None) is not None
            and m.Attachments.Count > 0,
        "categories": _parse_categories(getattr(m, "Categories", "") or ""),
        "importance": OL_IMPORTANCE_INV.get(getattr(m, "Importance", 1), "normal"),
        "folder": getattr(m.Parent, "Name", None) if getattr(m, "Parent", None) else None,
        "preview": preview,
    }


def _email_full(m) -> dict:
    base = _email_summary(m)
    attachments = []
    if m.Attachments and m.Attachments.Count:
        for i in range(1, m.Attachments.Count + 1):
            a = m.Attachments.Item(i)
            attachments.append({
                "filename": a.FileName,
                "size": getattr(a, "Size", None),
                "index": i,
            })
    base.update({
        "to": _recipients_to_list(m.Recipients),
        "body": m.Body or "",
        "body_html": getattr(m, "HTMLBody", None),
        "attachments": attachments,
        "conversation_id": getattr(m, "ConversationID", None),
        "conversation_topic": getattr(m, "ConversationTopic", None),
    })
    return base


def _event_summary(appt) -> dict:
    attendees = _recipients_to_list(appt.Recipients) if appt.Recipients else []
    return {
        "entry_id": appt.EntryID,
        "store_id": getattr(appt, "StoreID", None),
        "subject": appt.Subject,
        "start": _to_iso(appt.Start),
        "end": _to_iso(appt.End),
        "all_day": bool(getattr(appt, "AllDayEvent", False)),
        "location": getattr(appt, "Location", None),
        "organizer": getattr(appt, "Organizer", None),
        "attendees": attendees,
        "categories": _parse_categories(getattr(appt, "Categories", "") or ""),
        "is_recurring": bool(getattr(appt, "IsRecurring", False)),
        "is_meeting": getattr(appt, "MeetingStatus", 0) != 0,
        "response_status": getattr(appt, "ResponseStatus", None),
        "body": getattr(appt, "Body", None),
    }


# ---------- Accounts / folders ----------

def list_accounts() -> list[dict]:
    app, ns = _app_ns()
    out = []
    for i in range(1, ns.Accounts.Count + 1):
        acct = ns.Accounts.Item(i)
        out.append({
            "name": acct.DisplayName,
            "smtp": getattr(acct, "SmtpAddress", None),
            "type": getattr(acct, "AccountType", None),
        })
    return out


def list_folders(account: str | None = None, recursive: bool = True) -> list[dict]:
    _, ns = _app_ns()

    def walk(folder, depth=0) -> list[dict]:
        items = [{
            "name": folder.Name,
            "path": _folder_path(folder),
            "item_count": folder.Items.Count,
            "unread_count": getattr(folder, "UnReadItemCount", None),
            "depth": depth,
        }]
        if recursive and folder.Folders.Count:
            for i in range(1, folder.Folders.Count + 1):
                items.extend(walk(folder.Folders.Item(i), depth + 1))
        return items

    result = []
    stores = ns.Folders
    for i in range(1, stores.Count + 1):
        store = stores.Item(i)
        if account and store.Name.lower() != account.lower():
            continue
        result.extend(walk(store))
    return result


def _folder_path(folder) -> str:
    # Outlook's FolderPath is "\\Account\Inbox\Sub" — normalize to "Account/Inbox/Sub"
    fp = getattr(folder, "FolderPath", None)
    if fp:
        return fp.lstrip("\\").replace("\\", "/")
    return folder.Name


# ---------- Email ----------

def list_emails(
    folder: str | None = None,
    limit: int = 50,
    unread_only: bool = False,
    from_filter: str | None = None,
    subject_filter: str | None = None,
    since: str | None = None,
    before: str | None = None,
    newest_first: bool = True,
) -> list[dict]:
    f = _resolve_folder(folder)
    items = f.Items
    items.Sort("[ReceivedTime]", newest_first)

    restrict_parts = []
    if unread_only:
        restrict_parts.append("\"urn:schemas:httpmail:read\" = 0")
    if from_filter:
        esc = _escape_dasl(from_filter)
        restrict_parts.append(
            f"(\"urn:schemas:httpmail:fromemail\" LIKE '%{esc}%' OR "
            f"\"urn:schemas:httpmail:fromname\" LIKE '%{esc}%')"
        )
    if subject_filter:
        esc = _escape_dasl(subject_filter)
        restrict_parts.append(f"\"urn:schemas:httpmail:subject\" LIKE '%{esc}%'")
    if since:
        s = _format_filter_dt(_parse_dt(since))
        restrict_parts.append(f"\"urn:schemas:httpmail:datereceived\" >= '{s}'")
    if before:
        b = _format_filter_dt(_parse_dt(before))
        restrict_parts.append(f"\"urn:schemas:httpmail:datereceived\" < '{b}'")

    if restrict_parts:
        filt = "@SQL=" + " AND ".join(restrict_parts)
        try:
            items = items.Restrict(filt)
        except pywintypes.com_error as e:
            raise OutlookError(f"Filter failed: {filt!r}: {e}") from e

    out = []
    count = 0
    for item in items:
        if count >= limit:
            break
        # Skip non-mail items (calendar invites etc. in some folders)
        if getattr(item, "Class", 43) != 43:  # olMail = 43
            continue
        out.append(_email_summary(item))
        count += 1
    return out


def read_email(entry_id: str, store_id: str | None = None) -> dict:
    return _email_full(_get_item(entry_id, store_id))


def search_emails(
    query: str,
    folder: str | None = None,
    limit: int = 50,
) -> list[dict]:
    """Simple search: subject/sender/body LIKE %query%. Use list_emails with
    explicit filters for structured queries."""
    f = _resolve_folder(folder)
    items = f.Items
    items.Sort("[ReceivedTime]", True)
    esc = _escape_dasl(query)
    filt = (
        f"@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{esc}%' OR "
        f"\"urn:schemas:httpmail:fromname\" LIKE '%{esc}%' OR "
        f"\"urn:schemas:httpmail:textdescription\" LIKE '%{esc}%')"
    )
    try:
        items = items.Restrict(filt)
    except pywintypes.com_error as e:
        raise OutlookError(f"Search failed: {e}") from e
    out = []
    for item in items:
        if len(out) >= limit:
            break
        if getattr(item, "Class", 43) != 43:
            continue
        out.append(_email_summary(item))
    return out


def move_email(entry_id: str, target_folder: str, store_id: str | None = None) -> dict:
    item = _get_item(entry_id, store_id)
    target = _resolve_folder(target_folder)
    moved = item.Move(target)
    return {"entry_id": moved.EntryID, "store_id": moved.StoreID,
            "folder": target_folder}


def delete_email(entry_id: str, store_id: str | None = None,
                 permanent: bool = False) -> dict:
    item = _get_item(entry_id, store_id)
    if permanent:
        # Move to Deleted Items first, then delete from there to bypass restore
        _, ns = _app_ns()
        deleted = ns.GetDefaultFolder(OL_FOLDER_DELETED)
        item = item.Move(deleted)
        item.Delete()
        return {"deleted": True, "permanent": True}
    item.Delete()  # Standard Delete moves to Deleted Items
    return {"deleted": True, "permanent": False}


# ---------- Drafts / Send ----------

_REPLY_MODES = {"reply", "reply_all", "forward"}


def create_draft(
    to: Sequence[str] | None = None,
    cc: Sequence[str] | None = None,
    bcc: Sequence[str] | None = None,
    subject: str | None = None,
    body: str | None = None,
    html_body: str | None = None,
    attachments: Sequence[str] | None = None,
    reply_to_entry_id: str | None = None,
    reply_to_store_id: str | None = None,
    reply_mode: Literal["reply", "reply_all", "forward"] | None = None,
    importance: Literal["low", "normal", "high"] = "normal",
    categories: Sequence[str] | None = None,
) -> dict:
    app, _ = _app_ns()

    if reply_to_entry_id:
        if reply_mode not in _REPLY_MODES:
            raise OutlookError(
                f"reply_mode must be one of {_REPLY_MODES} when reply_to_entry_id is set"
            )
        source = _get_item(reply_to_entry_id, reply_to_store_id)
        if reply_mode == "reply":
            draft = source.Reply()
        elif reply_mode == "reply_all":
            draft = source.ReplyAll()
        else:
            draft = source.Forward()
    else:
        draft = app.CreateItem(OL_MAIL_ITEM)

    if to:
        draft.To = "; ".join(to)
    if cc:
        draft.CC = "; ".join(cc)
    if bcc:
        draft.BCC = "; ".join(bcc)
    if subject is not None:
        draft.Subject = subject
    if html_body is not None:
        draft.HTMLBody = _sanitize_html_body(html_body)
    elif body is not None:
        draft.Body = body
    if attachments:
        for path in attachments:
            _validate_attachment_path(path)
            draft.Attachments.Add(path)
    draft.Importance = OL_IMPORTANCE[importance]
    if categories is not None:
        draft.Categories = ",".join(categories)

    draft.Save()
    return _email_full(draft)


def update_draft(
    entry_id: str,
    store_id: str | None = None,
    to: Sequence[str] | None = None,
    cc: Sequence[str] | None = None,
    bcc: Sequence[str] | None = None,
    subject: str | None = None,
    body: str | None = None,
    html_body: str | None = None,
    add_attachments: Sequence[str] | None = None,
    importance: Literal["low", "normal", "high"] | None = None,
    categories: Sequence[str] | None = None,
) -> dict:
    draft = _get_item(entry_id, store_id)
    if to is not None:
        draft.To = "; ".join(to)
    if cc is not None:
        draft.CC = "; ".join(cc)
    if bcc is not None:
        draft.BCC = "; ".join(bcc)
    if subject is not None:
        draft.Subject = subject
    if html_body is not None:
        draft.HTMLBody = _sanitize_html_body(html_body)
    elif body is not None:
        draft.Body = body
    if add_attachments:
        for path in add_attachments:
            _validate_attachment_path(path)
            draft.Attachments.Add(path)
    if importance is not None:
        draft.Importance = OL_IMPORTANCE[importance]
    if categories is not None:
        draft.Categories = ",".join(categories)
    draft.Save()
    return _email_full(draft)


def send_email(
    entry_id: str | None = None,
    store_id: str | None = None,
    to: Sequence[str] | None = None,
    cc: Sequence[str] | None = None,
    bcc: Sequence[str] | None = None,
    subject: str | None = None,
    body: str | None = None,
    html_body: str | None = None,
    attachments: Sequence[str] | None = None,
    importance: Literal["low", "normal", "high"] = "normal",
) -> dict:
    if entry_id:
        draft = _get_item(entry_id, store_id)
    else:
        if not to or not subject:
            raise OutlookError("send_email requires either entry_id, or both to and subject")
        app, _ = _app_ns()
        draft = app.CreateItem(OL_MAIL_ITEM)
        draft.To = "; ".join(to)
        if cc:
            draft.CC = "; ".join(cc)
        if bcc:
            draft.BCC = "; ".join(bcc)
        draft.Subject = subject
        if html_body is not None:
            draft.HTMLBody = _sanitize_html_body(html_body)
        elif body is not None:
            draft.Body = body
        if attachments:
            for path in attachments:
                _validate_attachment_path(path)
                draft.Attachments.Add(path)
        draft.Importance = OL_IMPORTANCE[importance]

    # Capture before Send — after Send the COM reference becomes invalid
    # ("The item has been moved or deleted.")
    subject_sent = draft.Subject
    to_sent = draft.To
    draft.Send()
    return {"sent": True, "subject": subject_sent, "to": to_sent}


# ---------- Categories ----------

def list_categories() -> list[dict]:
    _, ns = _app_ns()
    cats = ns.Categories
    out = []
    for i in range(1, cats.Count + 1):
        c = cats.Item(i)
        out.append({
            "name": c.Name,
            "color": getattr(c, "Color", None),
            "shortcut_key": getattr(c, "ShortcutKey", None),
        })
    return out


def _apply_categories(item, categories: Sequence[str], mode: str):
    current = _parse_categories(getattr(item, "Categories", "") or "")
    if mode == "replace":
        new = list(categories)
    elif mode == "add":
        new = current + [c for c in categories if c not in current]
    elif mode == "remove":
        to_remove = set(categories)
        new = [c for c in current if c not in to_remove]
    else:
        raise OutlookError(f"categories mode must be replace/add/remove, got {mode!r}")
    item.Categories = ",".join(new)
    item.Save()
    return new


def set_email_categories(
    entry_id: str,
    categories: Sequence[str],
    mode: Literal["replace", "add", "remove"] = "replace",
    store_id: str | None = None,
) -> dict:
    item = _get_item(entry_id, store_id)
    new = _apply_categories(item, categories, mode)
    return {"entry_id": entry_id, "categories": new}


def set_event_categories(
    entry_id: str,
    categories: Sequence[str],
    mode: Literal["replace", "add", "remove"] = "replace",
    store_id: str | None = None,
) -> dict:
    item = _get_item(entry_id, store_id)
    new = _apply_categories(item, categories, mode)
    return {"entry_id": entry_id, "categories": new}


# ---------- Calendar ----------

def list_calendar_events(
    start: str,
    end: str,
    calendar: str | None = None,
    limit: int = 100,
) -> list[dict]:
    _, ns = _app_ns()
    if calendar:
        folder = _resolve_folder(calendar)
    else:
        folder = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
    items = folder.Items
    # Must set these for IncludeRecurrences to work correctly
    items.Sort("[Start]")
    items.IncludeRecurrences = True

    # Calendar uses Jet filters ([Start] >= '…'), which require the
    # Windows system-locale date format.  DASL URIs are not used here
    # because IncludeRecurrences only works with Jet Restrict.
    s = _format_jet_filter_dt(_parse_dt(start))
    e = _format_jet_filter_dt(_parse_dt(end))
    filt = f"[Start] >= '{s}' AND [Start] < '{e}'"
    try:
        restricted = items.Restrict(filt)
    except pywintypes.com_error as ex:
        raise OutlookError(f"Calendar filter failed: {ex}") from ex

    # With IncludeRecurrences=True the collection has no fixed Count,
    # so Python's ``for … in`` (which relies on _NewEnum / Count) may
    # silently yield nothing.  Use the GetFirst/GetNext COM pattern.
    out: list[dict] = []
    appt = restricted.GetFirst()
    while appt:
        if len(out) >= limit:
            break
        out.append(_event_summary(appt))
        appt = restricted.GetNext()
    return out


def create_calendar_event(
    subject: str,
    start: str,
    end: str,
    body: str | None = None,
    location: str | None = None,
    attendees: Sequence[str] | None = None,
    categories: Sequence[str] | None = None,
    all_day: bool = False,
    is_meeting: bool = False,
    reminder_minutes: int | None = None,
) -> dict:
    app, _ = _app_ns()
    appt = app.CreateItem(OL_APPT_ITEM)
    appt.Subject = subject
    appt.Start = _parse_dt(start)
    appt.End = _parse_dt(end)
    if all_day:
        appt.AllDayEvent = True
    if location:
        appt.Location = location
    if body is not None:
        appt.Body = body
    if categories:
        appt.Categories = ",".join(categories)
    if reminder_minutes is not None:
        appt.ReminderMinutesBeforeStart = reminder_minutes
        appt.ReminderSet = True
    if attendees:
        appt.MeetingStatus = 1  # olMeeting
        for a in attendees:
            appt.Recipients.Add(a)
        appt.Recipients.ResolveAll()
    elif is_meeting:
        appt.MeetingStatus = 1

    appt.Save()
    if is_meeting or attendees:
        appt.Send()
    return _event_summary(appt)


def update_calendar_event(
    entry_id: str,
    store_id: str | None = None,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    body: str | None = None,
    location: str | None = None,
    add_attendees: Sequence[str] | None = None,
    categories: Sequence[str] | None = None,
    send_update: bool = False,
) -> dict:
    appt = _get_item(entry_id, store_id)
    if subject is not None:
        appt.Subject = subject
    if start is not None:
        appt.Start = _parse_dt(start)
    if end is not None:
        appt.End = _parse_dt(end)
    if body is not None:
        appt.Body = body
    if location is not None:
        appt.Location = location
    if categories is not None:
        appt.Categories = ",".join(categories)
    if add_attendees:
        for a in add_attendees:
            appt.Recipients.Add(a)
        appt.Recipients.ResolveAll()
    appt.Save()
    if send_update and getattr(appt, "MeetingStatus", 0) != 0:
        appt.Send()
    return _event_summary(appt)


def delete_calendar_event(entry_id: str, store_id: str | None = None) -> dict:
    appt = _get_item(entry_id, store_id)
    appt.Delete()
    return {"deleted": True}


def respond_to_invite(
    entry_id: str,
    response: Literal["accept", "tentative", "decline"],
    send_response: bool = True,
    store_id: str | None = None,
) -> dict:
    if response not in OL_MEETING_RESPONSE:
        raise OutlookError(f"response must be accept/tentative/decline, got {response!r}")
    item = _get_item(entry_id, store_id)
    # MeetingItem → GetAssociatedAppointment; AppointmentItem → Respond directly
    appt = item.GetAssociatedAppointment(True) if hasattr(item, "GetAssociatedAppointment") else item
    resp = appt.Respond(OL_MEETING_RESPONSE[response], True, not send_response)
    if send_response:
        resp.Send()
    else:
        resp.Save()
    return {"responded": response, "sent": send_response}
