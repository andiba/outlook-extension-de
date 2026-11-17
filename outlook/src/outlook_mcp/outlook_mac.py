"""macOS backend for the Outlook MCP server.

Drives **classic** Microsoft Outlook for Mac (16.x) via its AppleScript
dictionary. Calls are dispatched through a single JXA (JavaScript for
Automation) script invoked via ``osascript``; payloads are JSON-encoded
in both directions, so the Python side stays small.

API mirrors :mod:`outlook_mcp.outlook_win`. ``store_id`` arguments are
accepted for source-compatibility but ignored — Mac AppleScript exposes
account scoping through the folder path instead.

**Limitations vs the Windows backend**

- ``id`` integers are stable as long as Outlook is running (and across
  restarts in practice), but Mac AppleScript does not expose a
  cross-machine identifier comparable to ``EntryID``.
- DASL-level full-text search is not available; ``search_emails``
  performs a JXA ``whose``-filter scan and may be slower than the
  Windows version on very large folders.
- ``list_emails`` and ``list_calendar_events`` select items through
  Spotlight (see "Spotlight fast path" below) and read only the hits via
  AppleScript: seconds instead of minutes/hours on large mailboxes.
  Recurring events are expanded from the RRULE Spotlight stores
  (python-dateutil); a changed occurrence is replaced by its exception
  record, an occurrence deleted from a series cannot be detected. Mail
  that arrived seconds ago may be missing until Spotlight indexed it.
  Without Spotlight the JXA path runs (bulk-fetches receive times; the
  calendar fallback does NOT expand recurrences — Outlook's dictionary
  has no ``expand`` for events).
- **Shared / delegated calendars are not readable via AppleScript.** The
  Spotlight path lists them by name (``calendar="Philipp Schulze"``);
  body and organizer may then be empty. Colleagues'
  calendars that you have permission to view show up in
  ``ol.calendars()`` (with name + account), but ``calendarEvents`` on
  them returns an empty list. Outlook for Mac fetches those events via
  EWS / Graph inside the app process and never materializes them
  through the AppleScript bridge. The Windows COM backend does not have
  this limitation. Reading shared calendars on Mac requires either the
  Outlook UI or a Graph integration (out of scope for this MCP).
- The "New Outlook for Mac" (Electron-based) has only partial
  AppleScript support — keep using the **classic** Outlook.
"""
from __future__ import annotations

import datetime as _dt
import json
import os
import plistlib
import re
import subprocess
from pathlib import Path
from typing import Any, Literal, Sequence

# Upper bound for one osascript call. A JXA scan of a very large folder can
# otherwise block the MCP server indefinitely.
_OSASCRIPT_TIMEOUT = int(os.environ.get("OUTLOOK_MCP_TIMEOUT", "180"))


class OutlookError(Exception):
    """User-facing error from the Outlook wrapper."""


# ---- security guards (mirrored from outlook_win.py) ----------------------

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
    return _DANGEROUS_HTML_RE.sub("<!-- blocked -->", html)


def _parse_iso(val: str | _dt.datetime) -> _dt.datetime:
    if isinstance(val, _dt.datetime):
        return val
    try:
        return _dt.datetime.fromisoformat(val)
    except ValueError as e:
        raise OutlookError(f"Invalid datetime (expected ISO 8601): {val!r}") from e


# ---- JXA dispatcher ------------------------------------------------------

# Localized folder aliases — superset of the Windows backend. Resolved on
# the Python side before being passed to JXA, so the JS knows only the
# canonical name (inbox/sent/drafts/deleted/junk/outbox).
_NAMED_FOLDERS = {
    "inbox": "inbox",
    "posteingang": "inbox",
    "sent": "sent",
    "sent items": "sent",
    "gesendete elemente": "sent",
    "gesendet": "sent",
    "drafts": "drafts",
    "entwürfe": "drafts",
    "deleted": "deleted",
    "deleted items": "deleted",
    "trash": "deleted",
    "gelöschte elemente": "deleted",
    "papierkorb": "deleted",
    "junk": "junk",
    "junk email": "junk",
    "junk-e-mail": "junk",
    "outbox": "outbox",
    "postausgang": "outbox",
}


_JXA_DISPATCH = r"""
function run(argv) {
  if (!argv || argv.length === 0) {
    return JSON.stringify({ ok: false, error: "missing payload" });
  }
  let payload;
  try { payload = JSON.parse(argv[0]); }
  catch (e) { return JSON.stringify({ ok: false, error: "bad JSON: " + e.message }); }
  const cmd = payload.cmd;
  const args = payload.args || {};
  const ol = Application("Microsoft Outlook");
  ol.includeStandardAdditions = true;
  try {
    const handler = HANDLERS[cmd];
    if (!handler) throw new Error("Unknown cmd: " + cmd);
    return JSON.stringify({ ok: true, result: handler(args, ol) });
  } catch (e) {
    return JSON.stringify({ ok: false, error: (e && e.message) ? e.message : String(e) });
  }
}

// ----- helpers ---------------------------------------------------------

function toISO(d) {
  if (d == null) return null;
  if (typeof d === "string") return d;
  try { return d.toISOString(); }
  catch (e) { return String(d); }
}

function safeGet(obj, prop) {
  try {
    const v = obj[prop]();
    return (v === undefined || v === null) ? null : v;
  } catch (e) { return null; }
}

function listToArray(list) {
  const out = [];
  if (!list) return out;
  let n = 0;
  try { n = list.length; } catch (e) { return out; }
  for (let i = 0; i < n; i++) {
    try { out.push(list[i]); } catch (e) { /* skip */ }
  }
  return out;
}

function findAccount(ol, name) {
  if (!name) return null;
  const lname = name.toLowerCase();
  const all = []
    .concat(listToArray(ol.exchangeAccounts()))
    .concat(listToArray(ol.imapAccounts()))
    .concat(listToArray(ol.popAccounts()));
  for (const acc of all) {
    const n = (safeGet(acc, "name") || "").toLowerCase();
    if (n === lname) return acc;
    const email = (safeGet(acc, "userName") || safeGet(acc, "emailAddress") || "").toLowerCase();
    if (email === lname) return acc;
  }
  return null;
}

function defaultAccount(ol) {
  // The application has a `default account` property (type: account).
  try { const a = ol.defaultAccount(); if (a) return a; } catch (e) {}
  // Fallback: first configured account by type.
  const ex = listToArray(ol.exchangeAccounts());
  if (ex.length) return ex[0];
  const imap = listToArray(ol.imapAccounts());
  if (imap.length) return imap[0];
  const pop = listToArray(ol.popAccounts());
  if (pop.length) return pop[0];
  return null;
}

// Map our canonical kind names to the actual JXA property names on the
// Outlook application/account class.
const FOLDER_PROP = {
  inbox:    "inbox",
  drafts:   "drafts",
  sent:     "sentItems",
  deleted:  "deletedItems",
  junk:     "junkMail",
  outbox:   "outbox",
};

function defaultStandardFolder(ol, kind) {
  const prop = FOLDER_PROP[kind];
  if (!prop) return null;
  // Try: account.<prop>(); fall back to the application-level relationship.
  const acc = defaultAccount(ol);
  if (acc) {
    try { const f = acc[prop](); if (f) return f; } catch (e) { /* fallthrough */ }
  }
  try { return ol[prop](); } catch (e) { return null; }
}

function findChildFolderByName(parent, name) {
  const ln = name.toLowerCase();
  const subs = listToArray(parent.mailFolders());
  for (const sub of subs) {
    const n = (safeGet(sub, "name") || "").toLowerCase();
    if (n === ln) return sub;
  }
  return null;
}

function walkSubfolders(folder, parts) {
  let cur = folder;
  for (const p of parts) {
    const next = findChildFolderByName(cur, p);
    if (!next) throw new Error("Subfolder not found: " + p);
    cur = next;
  }
  return cur;
}

// path: "{kind}/sub1/sub2" or "{accountName}/inbox/sub1" or "{kind}"
// kind is one of inbox|sent|drafts|deleted|junk|outbox (already canonicalized
// from the Python side via _NAMED_FOLDERS, or untouched if user passed an
// account name).
const STANDARD_KINDS = ["inbox","sent","drafts","deleted","junk","outbox"];

function resolveFolder(ol, path) {
  if (!path) {
    const f = defaultStandardFolder(ol, "inbox");
    if (!f) throw new Error("No default Inbox available");
    return f;
  }
  const parts = path.split("/").filter(p => p.length > 0);
  if (parts.length === 0) {
    const f = defaultStandardFolder(ol, "inbox");
    if (!f) throw new Error("No default Inbox available");
    return f;
  }
  const head = parts[0].toLowerCase();
  if (STANDARD_KINDS.indexOf(head) !== -1) {
    const root = defaultStandardFolder(ol, head);
    if (!root) throw new Error("Standard folder not available: " + head);
    return walkSubfolders(root, parts.slice(1));
  }
  // Try account name
  const acc = findAccount(ol, parts[0]);
  if (acc) {
    if (parts.length === 1) {
      try { return acc.inbox(); }
      catch (e) { throw new Error("Account has no inbox: " + parts[0]); }
    }
    const second = parts[1].toLowerCase();
    if (STANDARD_KINDS.indexOf(second) !== -1) {
      const prop = FOLDER_PROP[second];
      let root;
      try { root = acc[prop](); }
      catch (e) { throw new Error("Account has no " + second + ": " + parts[0]); }
      return walkSubfolders(root, parts.slice(2));
    }
    let inbox = null;
    try { inbox = acc.inbox(); } catch (e) { /* no inbox: top level only */ }
    const hit = walkFromParents(inbox ? [inbox] : [], parts.slice(1)) ||
                walkFromSiblings(listToArray(acc.mailFolders()), parts.slice(1));
    if (hit) return hit;
    throw new Error("Folder not found: " + path +
                    " (looked below the Inbox and among the account's top-level folders)");
  }
  // Fallback: below the default Inbox first (old behaviour), then among the
  // top-level folders — "HR/Urlaubsanträge" usually sits next to the Inbox,
  // not inside it.
  const inbox = defaultStandardFolder(ol, "inbox");
  const hit = walkFromParents(inbox ? [inbox] : [], parts) ||
              walkFromSiblings(topLevelFolders(ol), parts);
  if (hit) return hit;
  throw new Error("Folder not found: " + path +
                  " (looked below the Inbox and among the top-level folders)");
}

// First match of `parts` below any of `parents`, or null.
function walkFromParents(parents, parts) {
  for (const p of parents) {
    try { return walkSubfolders(p, parts); } catch (e) { /* next */ }
  }
  return null;
}

// First match where parts[0] is one of `folders` itself, or null.
function walkFromSiblings(folders, parts) {
  const ln = parts[0].toLowerCase();
  for (const f of folders) {
    if ((safeGet(f, "name") || "").toLowerCase() !== ln) continue;
    try { return walkSubfolders(f, parts.slice(1)); } catch (e) { /* next */ }
  }
  return null;
}

// Top-level mail folders: default account first, then the other accounts,
// then "On My Computer".
function topLevelFolders(ol) {
  const out = [];
  const def = defaultAccount(ol);
  if (def) { try { out.push(...listToArray(def.mailFolders())); } catch (e) {} }
  const accs = []
    .concat(listToArray(ol.exchangeAccounts()))
    .concat(listToArray(ol.imapAccounts()))
    .concat(listToArray(ol.popAccounts()));
  const defName = def ? safeGet(def, "name") : null;
  for (const a of accs) {
    if (defName && safeGet(a, "name") === defName) continue;
    try { out.push(...listToArray(a.mailFolders())); } catch (e) {}
  }
  try { out.push(...listToArray(ol.mailFolders())); } catch (e) {}
  return out;
}

function categoriesOf(item) {
  try {
    const cs = item.categories();
    return listToArray(cs).map(c => safeGet(c, "name")).filter(x => x);
  } catch (e) { return []; }
}

function categoryByName(ol, name) {
  const all = listToArray(ol.categories());
  const ln = name.toLowerCase();
  for (const c of all) {
    if ((safeGet(c, "name") || "").toLowerCase() === ln) return c;
  }
  return null;
}

function recipientsTo(arr, kind) {
  const out = [];
  for (const r of arr) {
    let name = null, email = null;
    try {
      const ea = r.emailAddress();
      if (ea) {
        try { name = ea.name; } catch (e) {}
        try { email = ea.address; } catch (e) {}
      }
    } catch (e) {}
    out.push({ name: name || null, email: email || null, type: kind });
  }
  return out;
}

function emailSummary(m) {
  const id = safeGet(m, "id");
  let preview = "";
  try {
    const body = safeGet(m, "plainTextContent") || safeGet(m, "content") || "";
    preview = String(body).slice(0, 200).replace(/\r\n/g, "\n").trim();
  } catch (e) { /* ignore */ }
  let folderName = null;
  try { folderName = safeGet(safeGet(m, "folder"), "name"); } catch (e) {}
  let senderName = null, senderEmail = null;
  try {
    const s = m.sender();
    if (s) {
      try { senderName = s.name || null; } catch (e) {}
      try { senderEmail = s.address || null; } catch (e) {}
    }
  } catch (e) {}
  let hasAttachments = false;
  try { hasAttachments = m.attachments().length > 0; } catch (e) {}
  let importance = "normal";
  try { importance = priorityName(safeGet(m, "priority")); } catch (e) {}
  return {
    entry_id: id != null ? String(id) : null,
    store_id: null,
    subject: safeGet(m, "subject"),
    sender_name: senderName,
    sender_email: senderEmail,
    received_time: toISO(safeGet(m, "timeReceived")),
    sent_time: toISO(safeGet(m, "timeSent")),
    unread: !safeGet(m, "isRead"),
    has_attachments: hasAttachments,
    categories: categoriesOf(m),
    importance: importance,
    folder: folderName,
    preview: preview,
  };
}

function emailFull(m) {
  const base = emailSummary(m);
  let body = "", htmlBody = "";
  try { body = safeGet(m, "plainTextContent") || safeGet(m, "content") || ""; } catch (e) {}
  try { htmlBody = safeGet(m, "htmlContent") || ""; } catch (e) {}
  const recipients = []
    .concat(recipientsTo(listToArray(m.toRecipients()), "to"))
    .concat(recipientsTo(listToArray(m.ccRecipients()), "cc"))
    .concat(recipientsTo(listToArray(m.bccRecipients()), "bcc"));
  const attachments = [];
  try {
    for (const a of listToArray(m.attachments())) {
      attachments.push({
        name: safeGet(a, "name"),
        size: safeGet(a, "fileSize"),
      });
    }
  } catch (e) {}
  base.body = body;
  base.html_body = htmlBody;
  base.recipients = recipients;
  base.attachments = attachments;
  return base;
}

function getMessageById(ol, id) {
  const numId = parseInt(id, 10);
  if (isNaN(numId)) throw new Error("Bad message id: " + id);
  // Outlook supports lookup by id directly via incoming/outgoing message classes
  try { return ol.incomingMessages.byId(numId); } catch (e) {}
  try { return ol.outgoingMessages.byId(numId); } catch (e) {}
  try { return ol.messages.byId(numId); } catch (e) {}
  throw new Error("Message not found: " + id);
}

function getEventById(ol, id) {
  const numId = parseInt(id, 10);
  if (isNaN(numId)) throw new Error("Bad event id: " + id);
  // Search across all calendars
  const cals = listToArray(ol.calendars());
  for (const cal of cals) {
    try {
      const ev = cal.calendarEvents.byId(numId);
      if (ev && ev.exists()) return ev;
    } catch (e) {}
  }
  try { return ol.calendarEvents.byId(numId); } catch (e) {}
  throw new Error("Calendar event not found: " + id);
}

function eventSummary(ev) {
  const id = safeGet(ev, "id");
  return {
    entry_id: id != null ? String(id) : null,
    store_id: null,
    subject: safeGet(ev, "subject"),
    location: safeGet(ev, "location"),
    start: toISO(safeGet(ev, "startTime")),
    end: toISO(safeGet(ev, "endTime")),
    all_day: !!safeGet(ev, "allDayFlag"),
    body: safeGet(ev, "plainTextContent") || safeGet(ev, "content") || "",
    categories: categoriesOf(ev),
    organizer: organizerOf(ev),
    is_recurring: !!safeGet(ev, "isRecurring"),
  };
}

function findCalendar(ol, name) {
  const cals = listToArray(ol.calendars());
  if (!name) {
    // Pick the calendar with the most events. The first calendar is
    // sometimes a placeholder with no name and an unusable element
    // accessor; the named "Calendar"/"Kalender" of the default account
    // is what the user thinks of as "their calendar".
    let best = null, bestCount = -1;
    for (const c of cals) {
      let cnt = -1;
      try { cnt = c.calendarEvents().length; } catch (e) { continue; }
      if (cnt > bestCount) { best = c; bestCount = cnt; }
    }
    if (best) return best;
    if (cals.length) return cals[0];
    throw new Error("No calendar available");
  }
  const ln = name.toLowerCase();
  for (const cal of cals) {
    if ((safeGet(cal, "name") || "").toLowerCase() === ln) return cal;
  }
  const acc = findAccount(ol, name);
  if (acc) {
    try { return acc.calendar(); } catch (e) {}
  }
  throw new Error("Calendar not found: " + name);
}

// ----- handlers --------------------------------------------------------

const HANDLERS = {

  list_accounts(args, ol) {
    const out = [];
    const groups = [
      ["exchange", listToArray(ol.exchangeAccounts())],
      ["imap",     listToArray(ol.imapAccounts())],
      ["pop",      listToArray(ol.popAccounts())],
    ];
    for (const [kind, list] of groups) {
      for (const acc of list) {
        out.push({
          name: safeGet(acc, "name"),
          email: safeGet(acc, "emailAddress") || safeGet(acc, "userName"),
          kind: kind,
          is_default: false,
        });
      }
    }
    const def = defaultAccount(ol);
    if (def) {
      const dn = safeGet(def, "name");
      for (const a of out) if (a.name === dn) a.is_default = true;
    }
    return out;
  },

  list_folders(args, ol) {
    // `path` is what folder= accepts: "HR/Urlaubsanträge".
    const collect = function(folder, depth, out, parent) {
      const name = safeGet(folder, "name");
      const path = parent ? parent + "/" + name : name;
      out.push({
        name: name,
        path: path,
        unread: safeGet(folder, "unreadCount"),
        total: safeGet(folder, "totalCount"),
        depth: depth,
      });
      if (args.recursive !== false) {
        for (const sub of listToArray(folder.mailFolders())) {
          collect(sub, depth + 1, out, path);
        }
      }
    };

    const out = [];
    if (args.account) {
      const acc = findAccount(ol, args.account);
      if (!acc) throw new Error("Account not found: " + args.account);
      for (const f of listToArray(acc.mailFolders())) collect(f, 0, out);
    } else {
      const accs = []
        .concat(listToArray(ol.exchangeAccounts()))
        .concat(listToArray(ol.imapAccounts()))
        .concat(listToArray(ol.popAccounts()));
      if (accs.length) {
        for (const acc of accs) {
          for (const f of listToArray(acc.mailFolders())) collect(f, 0, out);
        }
      } else {
        // No configured accounts — enumerate the application-level
        // mail folders (covers the "On My Computer" local store).
        for (const f of listToArray(ol.mailFolders())) collect(f, 0, out);
      }
    }
    return out;
  },

  list_emails(args, ol) {
    const folder = resolveFolder(ol, args.folder);
    const limit = args.limit || 50;
    const since = args.since ? new Date(args.since) : null;
    const before = args.before ? new Date(args.before) : null;
    // Fallback path (Spotlight unavailable). Fetch all receive times in ONE
    // Apple Event, filter and sort in JS, then touch only the hits. The old
    // code sorted message references and read timeReceived per comparison
    // (~2·n·log n Apple Events at ~16 ms each: hours on a 36k inbox).
    const spec = folder.messages;
    let times = [];
    try { times = spec.timeReceived(); } catch (e) { times = []; }
    let idx = [];
    for (let i = 0; i < times.length; i++) {
      const t = times[i];
      if (since && t && t < since) continue;
      if (before && t && t >= before) continue;
      idx.push(i);
    }
    idx.sort(function(a, b) {
      const va = times[a] ? times[a].getTime() : 0;
      const vb = times[b] ? times[b].getTime() : 0;
      return args.newest_first === false ? va - vb : vb - va;
    });
    const msgs = [];
    for (const i of idx) { try { msgs.push(spec[i]); } catch (e) {} }
    const fromQ = args.from_filter ? args.from_filter.toLowerCase() : null;
    const subjQ = args.subject_filter ? args.subject_filter.toLowerCase() : null;
    const out = [];
    for (const m of msgs) {
      if (out.length >= limit) break;
      if (args.unread_only && safeGet(m, "isRead")) continue;
      if (subjQ) {
        const s = (safeGet(m, "subject") || "").toLowerCase();
        if (s.indexOf(subjQ) === -1) continue;
      }
      if (fromQ) {
        let sn = "", sa = "";
        try {
          const s = m.sender();
          if (s) {
            try { sn = (s.name || "").toLowerCase(); } catch (e) {}
            try { sa = (s.address || "").toLowerCase(); } catch (e) {}
          }
        } catch (e) {}
        if (sn.indexOf(fromQ) === -1 && sa.indexOf(fromQ) === -1) continue;
      }
      out.push(emailSummary(m));
    }
    return out;
  },

  read_email(args, ol) {
    const m = getMessageById(ol, args.entry_id);
    return emailFull(m);
  },

  search_emails(args, ol) {
    const q = (args.query || "").toLowerCase();
    if (!q) return [];
    const folder = resolveFolder(ol, args.folder);
    const limit = args.limit || 50;
    const msgs = listToArray(folder.messages());
    const out = [];
    for (const m of msgs) {
      if (out.length >= limit) break;
      const subj = (safeGet(m, "subject") || "").toLowerCase();
      let body = "";
      try { body = (safeGet(m, "plainTextContent") || "").toLowerCase(); } catch (e) {}
      let senderName = "", senderAddr = "";
      try {
        const s = m.sender();
        if (s) {
          try { senderName = (s.name || "").toLowerCase(); } catch (e) {}
          try { senderAddr = (s.address || "").toLowerCase(); } catch (e) {}
        }
      } catch (e) {}
      if (subj.indexOf(q) !== -1 ||
          body.indexOf(q) !== -1 ||
          senderName.indexOf(q) !== -1 ||
          senderAddr.indexOf(q) !== -1) {
        out.push(emailSummary(m));
      }
    }
    return out;
  },

  move_email(args, ol) {
    const m = getMessageById(ol, args.entry_id);
    const target = resolveFolder(ol, args.target_folder);
    ol.move(m, { to: target });
    return { ok: true, target_folder: args.target_folder };
  },

  delete_email(args, ol) {
    const m = getMessageById(ol, args.entry_id);
    if (args.permanent) {
      ol.permanentlyDelete(m);
    } else {
      ol.delete(m);
    }
    return { ok: true, permanent: !!args.permanent };
  },

  // create_draft is split: JXA creates the base outgoing message and sets
  // subject/body/importance/categories; the Python wrapper layers the
  // recipients on via separate AppleScript invocations (JXA + NSAppleScript
  // co-resident has crashed osascript on this Outlook build).
  create_draft(args, ol) {
    let draft;
    if (args.reply_to_entry_id) {
      const orig = getMessageById(ol, args.reply_to_entry_id);
      const mode = args.reply_mode || "reply";
      if (mode === "reply") draft = ol.reply(orig);
      else if (mode === "reply_all") draft = ol.replyAll(orig);
      else if (mode === "forward") draft = ol.forward(orig);
      else throw new Error("Unknown reply_mode: " + mode);
      if (args.body || args.html_body) {
        const newText = args.body || "";
        const oldText = safeGet(draft, "plainTextContent") || "";
        draft.plainTextContent = newText + "\n\n" + oldText;
      }
    } else {
      const props = {
        subject: args.subject || "",
        plainTextContent: args.body || "",
      };
      if (args.html_body) props.content = args.html_body;
      draft = ol.OutgoingMessage(props);
      draft.make();
    }
    if (args.subject && !args.reply_to_entry_id) draft.subject = args.subject;
    if (args.attachments) {
      for (const path of args.attachments) {
        ol.Attachment({ file: Path(path) }).make({ at: draft });
      }
    }
    if (args.importance && args.importance !== "normal") draft.priority = args.importance;
    if (args.categories) {
      const cats = [];
      for (const cn of args.categories) {
        const c = categoryByName(ol, cn);
        if (c) cats.push(c);
      }
      if (cats.length) draft.categories = cats;
    }
    return emailSummary(draft);
  },

  update_draft(args, ol) {
    const m = getMessageById(ol, args.entry_id);
    if (args.subject != null) m.subject = args.subject;
    if (args.body != null) m.plainTextContent = args.body;
    if (args.html_body != null) m.content = args.html_body;
    if (args.importance != null) m.priority = args.importance;
    if (args.add_attachments) {
      for (const path of args.add_attachments) {
        ol.Attachment({ file: Path(path) }).make({ at: m });
      }
    }
    if (args.categories) {
      const cats = [];
      for (const cn of args.categories) {
        const c = categoryByName(ol, cn);
        if (c) cats.push(c);
      }
      m.categories = cats;
    }
    // Note: don't call m.save() — in Outlook for Mac AppleScript, `save`
    // means "save to file"; property assignments auto-persist.
    return emailSummary(m);
  },

  send_email(args, ol) {
    let m;
    if (args.entry_id) {
      m = getMessageById(ol, args.entry_id);
    } else {
      m = HANDLERS.create_draft(args, ol);
      m = getMessageById(ol, m.entry_id);
    }
    ol.send(m);
    return { ok: true, sent_at: new Date().toISOString() };
  },

  list_categories(args, ol) {
    const out = [];
    for (const c of listToArray(ol.categories())) {
      out.push({
        name: safeGet(c, "name"),
        color: safeGet(c, "color"),
        shortcut: null,
      });
    }
    return out;
  },

  set_email_categories(args, ol) {
    const m = getMessageById(ol, args.entry_id);
    return _setCategories(ol, m, args);
  },

  set_event_categories(args, ol) {
    const ev = getEventById(ol, args.entry_id);
    return _setCategories(ol, ev, args);
  },

  list_calendar_events(args, ol) {
    const cal = findCalendar(ol, args.calendar);
    const start = new Date(args.start);
    const end = new Date(args.end);
    const limit = args.limit || 100;
    let events = listToArray(cal.calendarEvents());
    // Expand recurring events that overlap the window
    const expanded = [];
    for (const ev of events) {
      try {
        const evStart = safeGet(ev, "startTime");
        const evEnd = safeGet(ev, "endTime");
        const isRec = !!safeGet(ev, "isRecurring");
        if (isRec) {
          // Use AppleScript "expand" to materialize occurrences in the window
          let occs = [];
          try {
            occs = ol.expand(ev, { from: start, until: end }) || [];
          } catch (e) {}
          for (const occ of listToArray(occs)) expanded.push(occ);
        } else {
          if (evEnd && evEnd <= start) continue;
          if (evStart && evStart >= end) continue;
          expanded.push(ev);
        }
      } catch (e) {}
    }
    expanded.sort(function(a, b) {
      const ta = safeGet(a, "startTime");
      const tb = safeGet(b, "startTime");
      return (ta ? ta.getTime() : 0) - (tb ? tb.getTime() : 0);
    });
    return expanded.slice(0, limit).map(eventSummary);
  },

  create_calendar_event(args, ol) {
    const cal = findCalendar(ol, null);
    const props = {
      subject: args.subject,
      startTime: new Date(args.start),
      endTime: new Date(args.end),
      allDayFlag: !!args.all_day,
    };
    if (args.body) props.plainTextContent = args.body;
    if (args.location) props.location = args.location;
    if (args.reminder_minutes != null) {
      props.reminderTime = args.reminder_minutes;
    }
    const ev = ol.CalendarEvent(props);
    ev.make({ at: cal });
    if ((args.attendees && args.attendees.length) || args.is_meeting) {
      try { ol.send(ev); } catch (e) { /* invitation send is best-effort */ }
    }
    return eventSummary(ev);
  },

  update_calendar_event(args, ol) {
    const ev = getEventById(ol, args.entry_id);
    if (args.subject != null) ev.subject = args.subject;
    if (args.start != null) ev.startTime = new Date(args.start);
    if (args.end != null) ev.endTime = new Date(args.end);
    if (args.body != null) ev.plainTextContent = args.body;
    if (args.location != null) ev.location = args.location;
    if (args.send_update) {
      try { ol.send(ev); } catch (e) {}
    }
    return eventSummary(ev);
  },

  delete_calendar_event(args, ol) {
    const ev = getEventById(ol, args.entry_id);
    ol.delete(ev);
    return { ok: true };
  },

  respond_to_invite(args, ol) {
    const m = getMessageById(ol, args.entry_id);
    const send = args.send_response !== false;
    const r = args.response;
    if (r === "accept") {
      send ? ol.acceptMeeting(m) : ol.acceptMeeting(m, { withResponse: false });
    } else if (r === "tentative") {
      send ? ol.acceptTentativelyMeeting(m) : ol.acceptTentativelyMeeting(m, { withResponse: false });
    } else if (r === "decline") {
      send ? ol.declineMeeting(m) : ol.declineMeeting(m, { withResponse: false });
    } else {
      throw new Error("Bad response: " + r);
    }
    return { ok: true, response: r, sent: send };
  },

  // ---- helpers for the Spotlight fast path (see _spotlight_* in Python) --

  spotlight_mail_context(args, ol) {
    const f = resolveFolder(ol, args.folder);
    let count = -1;
    try { count = f.messages.length; } catch (e) {}
    const cats = {};
    try {
      const c = ol.categories;
      const ids = c.id(), names = c.name();
      for (let i = 0; i < ids.length; i++) cats[String(ids[i])] = names[i];
    } catch (e) {}
    return { id: safeGet(f, "id"), name: safeGet(f, "name"), count: count, categories: cats };
  },

  spotlight_calendar_context(args, ol) {
    const out = { calendars: [], categories: {} };
    try {
      const c = ol.calendars;
      const ids = c.id(), names = c.name();
      for (let i = 0; i < ids.length; i++) out.calendars.push({ id: ids[i], name: names[i] });
    } catch (e) {}
    try {
      const c = ol.categories;
      const ids = c.id(), names = c.name();
      for (let i = 0; i < ids.length; i++) out.categories[String(ids[i])] = names[i];
    } catch (e) {}
    return out;
  },

  email_extras(args, ol) {
    const out = {};
    for (const id of (args.ids || [])) {
      let m;
      try { m = ol.messages.byId(parseInt(id, 10)); m.subject(); }
      catch (e) { try { m = getMessageById(ol, String(id)); } catch (e2) { continue; } }
      const p = safeGet(m, "priority");
      let preview = "";
      if (args.preview !== false) {
        try {
          const body = safeGet(m, "plainTextContent") || "";
          preview = String(body).slice(0, 200).replace(/\r\n/g, "\n").trim();
        } catch (e) {}
      }
      out[String(id)] = { importance: priorityName(p), preview: preview };
    }
    return out;
  },

  event_extras(args, ol) {
    const out = {};
    for (const id of (args.ids || [])) {
      let ev = null;
      try { ev = ol.calendarEvents.byId(parseInt(id, 10)); ev.subject(); }
      catch (e) { ev = null; }
      if (!ev) continue;
      out[String(id)] = {
        location: safeGet(ev, "location"),
        organizer: organizerOf(ev),
        body: args.body === false ? "" : (safeGet(ev, "plainTextContent") || ""),
      };
    }
    return out;
  },
};

function organizerOf(ev) {
  // Outlook returns the organizer as a plain address string; older code
  // expected an object with .address and therefore always got null.
  const o = safeGet(ev, "organizer");
  if (o == null) return null;
  if (typeof o === "string") return o || null;
  try { return o.address || null; } catch (e) { return null; }
}

function priorityName(p) {
  if (p === "priority high" || p === "high") return "high";
  if (p === "priority low" || p === "low") return "low";
  return "normal";
}

function _setCategories(ol, item, args) {
  const requested = [];
  for (const cn of (args.categories || [])) {
    const c = categoryByName(ol, cn);
    if (c) requested.push(c);
  }
  const mode = args.mode || "replace";
  if (mode === "replace") {
    item.categories = requested;
  } else if (mode === "add") {
    const cur = listToArray(item.categories());
    const seen = {};
    for (const c of cur) seen[safeGet(c, "name")] = true;
    const merged = cur.slice();
    for (const c of requested) if (!seen[safeGet(c, "name")]) merged.push(c);
    item.categories = merged;
  } else if (mode === "remove") {
    const remove = {};
    for (const c of requested) remove[safeGet(c, "name")] = true;
    const cur = listToArray(item.categories());
    item.categories = cur.filter(c => !remove[safeGet(c, "name")]);
  } else {
    throw new Error("Bad mode: " + mode);
  }
  return { ok: true, mode: mode, categories: categoriesOf(item) };
}
"""


def _canonical_folder_path(path: str | None) -> str | None:
    """Canonicalize the leading folder segment via _NAMED_FOLDERS aliases.

    The JXA side only knows the canonical English folder kinds. Account
    names and unknown folder names pass through untouched so subfolder
    lookups still work.
    """
    if not path:
        return path
    parts = [p for p in path.split("/") if p]
    if not parts:
        return path
    head = parts[0].lower()
    if head in _NAMED_FOLDERS:
        parts[0] = _NAMED_FOLDERS[head]
    return "/".join(parts)


def _as_escape(s: str | None) -> str:
    """Escape a string for embedding inside an AppleScript double-quoted literal."""
    if s is None:
        return ""
    return str(s).replace("\\", "\\\\").replace('"', '\\"')


def _run_applescript(src: str) -> str:
    """Run an AppleScript snippet via osascript -e (separate process).

    Used for operations that crash JXA — notably ``make new recipient``
    and meeting-attendee creation. Each call is its own process so JXA
    state never co-exists with NSAppleScript state.
    """
    try:
        proc = subprocess.run(
            ["osascript", "-e", src],
            capture_output=True, text=True, check=False,
            timeout=_OSASCRIPT_TIMEOUT,
        )
    except subprocess.TimeoutExpired as e:
        raise OutlookError(
            f"AppleScript timed out after {_OSASCRIPT_TIMEOUT}s"
        ) from e
    if proc.returncode != 0:
        stderr = (proc.stderr or "").strip()
        raise OutlookError(f"AppleScript failed: {stderr or proc.stdout!r}")
    return (proc.stdout or "").strip()


_RECIP_CLASS = {"to": "to recipient", "cc": "cc recipient", "bcc": "bcc recipient"}
_RECIP_ELEM = {"to": "to recipient", "cc": "cc recipient", "bcc": "bcc recipient"}


def _create_draft_as(
    *,
    subject: str | None,
    body: str | None,
    html_body: str | None,
    importance: str,
    to: Sequence[str] | None,
    cc: Sequence[str] | None,
    bcc: Sequence[str] | None,
    reply_to_entry_id: str | None,
    reply_mode: str | None,
    categories: Sequence[str] | None,
    attachments: Sequence[str] | None,
) -> str:
    """Create an outgoing message via pure AppleScript and return its id.

    Pure-AS keeps all object handles inside one Outlook event chain and
    avoids the AppleEvent -10000 errors triggered by mixing JXA and AS on
    the same draft.
    """
    lines = ['tell application "Microsoft Outlook"']
    if reply_to_entry_id:
        cmd = {"reply": "reply", "reply_all": "reply all", "forward": "forward"}.get(
            reply_mode or "reply", "reply",
        )
        lines.append(f'set orig to incoming message id {int(reply_to_entry_id)}')
        lines.append(f'set m to {cmd} orig')
        if subject:
            lines.append(f'set subject of m to "{_as_escape(subject)}"')
        if body:
            old_body_ref = "plain text content of m"
            lines.append(
                f'set plain text content of m to '
                f'"{_as_escape(body)}" & return & return & ({old_body_ref})'
            )
        if html_body:
            lines.append(f'set content of m to "{_as_escape(html_body)}"')
    else:
        props = []
        if subject is not None: props.append(f'subject:"{_as_escape(subject)}"')
        if body is not None: props.append(f'plain text content:"{_as_escape(body)}"')
        if html_body is not None: props.append(f'content:"{_as_escape(html_body)}"')
        if importance and importance != "normal":
            props.append(f'priority:{importance}')
        prop_clause = "{" + ", ".join(props) + "}" if props else "{}"
        lines.append(f'set m to make new outgoing message with properties {prop_clause}')

    for addr in (to or []):
        lines.append(
            f'make new to recipient at m with properties '
            f'{{email address:{{address:"{_as_escape(addr)}"}}}}'
        )
    for addr in (cc or []):
        lines.append(
            f'make new cc recipient at m with properties '
            f'{{email address:{{address:"{_as_escape(addr)}"}}}}'
        )
    for addr in (bcc or []):
        lines.append(
            f'make new bcc recipient at m with properties '
            f'{{email address:{{address:"{_as_escape(addr)}"}}}}'
        )

    if categories:
        cat_refs = ", ".join(
            f'category "{_as_escape(c)}"' for c in categories
        )
        lines.append(f'set categories of m to {{{cat_refs}}}')

    if attachments:
        for path in attachments:
            lines.append(
                f'make new attachment at m with properties '
                f'{{file:POSIX file "{_as_escape(path)}"}}'
            )

    lines.append('return id of m as text')
    lines.append('end tell')
    return _run_applescript("\n".join(lines))


def _set_categories_as(
    entry_id: str,
    categories: Sequence[str],
    mode: Literal["replace", "add", "remove"],
    item_class: Literal["message", "event"],
) -> dict:
    """Set categories on a message or calendar event via AppleScript.

    JXA category assignment fails on outgoing messages (and possibly
    other Outlook-Mac items) with "types can't be converted"; AS handles
    it cleanly via the named ``category "Name"`` reference.
    """
    item_phrase = (
        "(first message whose id is " if item_class == "message"
        else "(first calendar event whose id is "
    )
    add_lines: list[str] = []
    if mode == "replace":
        cat_refs = ", ".join(f'category "{_as_escape(c)}"' for c in categories)
        add_lines.append(
            f'set categories of m to {{{cat_refs}}}' if categories else
            'set categories of m to {}'
        )
    elif mode == "add":
        add_lines.append('set existingNames to {}')
        add_lines.append('set existingCats to categories of m')
        add_lines.append('repeat with i from 1 to count of existingCats')
        add_lines.append('  set end of existingNames to name of item i of existingCats')
        add_lines.append('end repeat')
        add_lines.append('set newCats to {}')
        add_lines.append('repeat with i from 1 to count of existingCats')
        add_lines.append('  set end of newCats to item i of existingCats')
        add_lines.append('end repeat')
        for c in categories:
            add_lines.append(
                f'if "{_as_escape(c)}" is not in existingNames then '
                f'set end of newCats to category "{_as_escape(c)}"'
            )
        add_lines.append('set categories of m to newCats')
    elif mode == "remove":
        add_lines.append('set newCats to {}')
        add_lines.append('set existingCats to categories of m')
        add_lines.append('repeat with i from 1 to count of existingCats')
        keep_clauses = " and ".join(
            f'name of item i of existingCats is not "{_as_escape(name)}"' for name in categories
        ) or "true"
        add_lines.append(f'  if {keep_clauses} then set end of newCats to item i of existingCats')
        add_lines.append('end repeat')
        add_lines.append('set categories of m to newCats')
    else:
        raise OutlookError(f"Bad mode: {mode}")

    src = (
        'tell application "Microsoft Outlook"\n'
        f'set m to {item_phrase}{int(entry_id)})\n'
        + "\n".join(add_lines) + "\n"
        'set out to ""\n'
        'set catlist to categories of m\n'
        'repeat with i from 1 to count of catlist\n'
        '  set out to out & (name of item i of catlist) & "|"\n'
        'end repeat\n'
        'return out\n'
        'end tell'
    )
    raw = _run_applescript(src)
    cats = [c for c in raw.split("|") if c]
    return {"ok": True, "mode": mode, "categories": cats}


def _summarize_message_as(entry_id: str) -> dict:
    """Build a basic email summary via AppleScript only.

    Used by ``create_draft`` and ``update_draft`` so the wrapper never
    touches the draft via JXA — JXA reads have been observed to put
    Outlook for Mac 16.x into a state where subsequent AS modifications
    on the same draft fail with AppleEvent error -10000.

    Output mirrors the JXA ``read_email`` shape (minus body details that
    aren't needed for confirmation responses).
    """
    src = f'''
set q to "\\""
set lf to ASCII character 10
tell application "Microsoft Outlook"
  set m to outgoing message id {int(entry_id)}
  set ssubj to subject of m
  set sbody to plain text content of m
  set spri to priority of m as text
  set sread to is read of m as text
  set scats to ""
  set catlist to categories of m
  repeat with i from 1 to count of catlist
    set scats to scats & (name of item i of catlist) & "|"
  end repeat
  set rcps to ""
  repeat with r in to recipients of m
    set ea to email address of r
    try
      set rname to name of ea
    on error
      set rname to ""
    end try
    set rcps to rcps & "to||" & rname & "||" & address of ea & "<<<"
  end repeat
  repeat with r in cc recipients of m
    set ea to email address of r
    try
      set rname to name of ea
    on error
      set rname to ""
    end try
    set rcps to rcps & "cc||" & rname & "||" & address of ea & "<<<"
  end repeat
  repeat with r in bcc recipients of m
    set ea to email address of r
    try
      set rname to name of ea
    on error
      set rname to ""
    end try
    set rcps to rcps & "bcc||" & rname & "||" & address of ea & "<<<"
  end repeat
  return ssubj & "<FLD>" & sbody & "<FLD>" & spri & "<FLD>" & sread & "<FLD>" & scats & "<FLD>" & rcps
end tell
'''
    raw = _run_applescript(src)
    parts = raw.split("<FLD>", 5)
    while len(parts) < 6:
        parts.append("")
    subject, body, priority, is_read, cats_raw, rcps_raw = parts
    cats = [c for c in cats_raw.split("|") if c]
    recipients: list[dict] = []
    for chunk in rcps_raw.split("<<<"):
        if not chunk:
            continue
        kind, name, address = (chunk.split("||", 2) + ["", "", ""])[:3]
        recipients.append({"name": name or None, "email": address or None, "type": kind})
    return {
        "entry_id": str(int(entry_id)),
        "store_id": None,
        "subject": subject or None,
        "sender_name": None,
        "sender_email": None,
        "received_time": None,
        "sent_time": None,
        "unread": is_read.strip().lower() != "true",
        "has_attachments": False,
        "categories": cats,
        "importance": priority if priority in ("low", "normal", "high") else "normal",
        "folder": "drafts",
        "preview": (body or "")[:200].replace("\r\n", "\n").strip(),
        "body": body or "",
        "html_body": "",
        "recipients": recipients,
        "attachments": [],
    }


def _update_draft_as(
    entry_id: str,
    *,
    subject: str | None,
    body: str | None,
    html_body: str | None,
    importance: str | None,
    to: Sequence[str] | None,
    cc: Sequence[str] | None,
    bcc: Sequence[str] | None,
    add_attachments: Sequence[str] | None,
    categories: Sequence[str] | None,
) -> None:
    """Update an existing draft entirely via AppleScript.

    Pure-AS sidesteps the JXA-vs-AS interaction bug observed when
    flipping recipient lists on Outlook for Mac 16.x.
    """
    lines = [
        'tell application "Microsoft Outlook"',
        f'set m to outgoing message id {int(entry_id)}',
    ]
    if subject is not None:
        lines.append(f'set subject of m to "{_as_escape(subject)}"')
    if body is not None:
        lines.append(f'set plain text content of m to "{_as_escape(body)}"')
    if html_body is not None:
        lines.append(f'set content of m to "{_as_escape(html_body)}"')
    if importance:
        lines.append(f'set priority of m to {importance}')

    # Recipients: pure AS lets delete + make on the same draft within one
    # tell block work reliably; the JXA-induced state is not present.
    for kind, addrs in (("to", to), ("cc", cc), ("bcc", bcc)):
        if addrs is None:
            continue
        elem = _RECIP_ELEM[kind]
        cls = _RECIP_CLASS[kind]
        lines.append(f'delete every {elem} of m')
        for addr in addrs:
            lines.append(
                f'make new {cls} at m with properties '
                f'{{email address:{{address:"{_as_escape(addr)}"}}}}'
            )

    if add_attachments:
        for path in add_attachments:
            lines.append(
                f'make new attachment at m with properties '
                f'{{file:POSIX file "{_as_escape(path)}"}}'
            )

    if categories is not None:
        if categories:
            cat_refs = ", ".join(
                f'category "{_as_escape(c)}"' for c in categories
            )
            lines.append(f'set categories of m to {{{cat_refs}}}')
        else:
            lines.append('set categories of m to {}')

    lines.append('end tell')
    _run_applescript("\n".join(lines))


def _add_attendee(event_id: str, address: str, optional: bool = False) -> None:
    cls = "optional attendee" if optional else "required attendee"
    src = (
        'tell application "Microsoft Outlook"\n'
        f'set ev to calendar event id {int(event_id)}\n'
        f'make new {cls} at ev with properties {{email address:{{address:"{_as_escape(address)}"}}}}\n'
        'end tell'
    )
    _run_applescript(src)


def _run_jxa(cmd: str, args: dict | None = None) -> Any:
    payload = json.dumps({"cmd": cmd, "args": args or {}})
    try:
        proc = subprocess.run(
            ["osascript", "-l", "JavaScript", "-e", _JXA_DISPATCH, payload],
            capture_output=True, text=True, check=False,
            timeout=_OSASCRIPT_TIMEOUT,
        )
    except FileNotFoundError as e:
        raise OutlookError("osascript not found — macOS only backend") from e
    except subprocess.TimeoutExpired as e:
        raise OutlookError(
            f"Outlook did not answer within {_OSASCRIPT_TIMEOUT}s ({cmd}). "
            "Set OUTLOOK_MCP_TIMEOUT to allow more time."
        ) from e
    if proc.returncode != 0:
        stderr = (proc.stderr or "").strip()
        if "execution error" in stderr.lower() and "-1743" in stderr:
            raise OutlookError(
                "macOS Automation permission denied. Open System Settings → "
                "Privacy & Security → Automation, find the calling app, and "
                "enable 'Microsoft Outlook'."
            )
        raise OutlookError(f"osascript failed: {stderr or proc.stdout!r}")
    out = (proc.stdout or "").strip()
    if not out:
        raise OutlookError("osascript returned empty output")
    try:
        envelope = json.loads(out)
    except json.JSONDecodeError as e:
        raise OutlookError(f"osascript returned non-JSON output: {out!r}") from e
    if not envelope.get("ok"):
        raise OutlookError(envelope.get("error", "unknown JXA error"))
    return envelope.get("result")


# ---- Spotlight fast path -------------------------------------------------
#
# Classic Outlook for Mac keeps every message and event as a file in its
# profile, and Spotlight indexes their metadata: record id, folder id,
# dates, sender, unread flag, recurrence rule. One ``mdfind`` answers
# "what arrived in folder Y since X" in well under a second. AppleScript
# needs one Apple Event (~16 ms) per item AND property, which made
# list_emails/list_calendar_events take minutes to hours on large
# mailboxes. AppleScript is now used only for the few hits afterwards
# (preview, priority, organizer, body). If Spotlight is unavailable (index
# off, profile elsewhere, OUTLOOK_MCP_SPOTLIGHT=0) the JXA path runs.

_OUTLOOK_DATA = Path(os.environ.get(
    "OUTLOOK_MCP_PROFILE_DIR",
    str(Path.home() / "Library/Group Containers/UBF8T346G9.Office/Outlook"),
))
_MD_MESSAGE = "com.microsoft.outlook15.message"
_MD_EVENT = "com.microsoft.outlook15.event"
_MDP = "com_microsoft_outlook_"
_SPOTLIGHT_TIMEOUT = 30
_UTC = _dt.timezone.utc


class _SpotlightUnavailable(Exception):
    """Spotlight cannot answer — fall back to the JXA path."""


def _spotlight_enabled() -> bool:
    return (os.environ.get("OUTLOOK_MCP_SPOTLIGHT", "1") != "0"
            and _OUTLOOK_DATA.is_dir())


def _to_utc(dt: _dt.datetime) -> _dt.datetime:
    # A naive value is local wall time, exactly as JS `new Date(iso)` read it.
    if dt.tzinfo is None:
        dt = dt.astimezone()
    return dt.astimezone(_UTC)


def _md_time(dt: _dt.datetime) -> str:
    return "$time.iso(" + _to_utc(dt).strftime("%Y-%m-%dT%H:%M:%SZ") + ")"


def _md_iso(val: Any) -> str | None:
    """plistlib dates are naive UTC; format like JS toISOString()."""
    if not isinstance(val, _dt.datetime):
        return None
    if val.tzinfo is not None:
        val = val.astimezone(_UTC).replace(tzinfo=None)
    return val.strftime("%Y-%m-%dT%H:%M:%S.000Z")


def _md_utc(val: Any) -> _dt.datetime | None:
    if not isinstance(val, _dt.datetime):
        return None
    return val.replace(tzinfo=_UTC) if val.tzinfo is None else val.astimezone(_UTC)


def _mdfind(query: str) -> list[str]:
    try:
        proc = subprocess.run(
            ["mdfind", "-onlyin", str(_OUTLOOK_DATA), query],
            capture_output=True, text=True, check=False,
            timeout=_SPOTLIGHT_TIMEOUT,
        )
    except (FileNotFoundError, subprocess.TimeoutExpired) as e:
        raise _SpotlightUnavailable(str(e)) from e
    if proc.returncode != 0:
        raise _SpotlightUnavailable((proc.stderr or "").strip())
    return [p for p in proc.stdout.splitlines() if p]


def _mdfind_count(query: str) -> int:
    try:
        proc = subprocess.run(
            ["mdfind", "-count", "-onlyin", str(_OUTLOOK_DATA), query],
            capture_output=True, text=True, check=False,
            timeout=_SPOTLIGHT_TIMEOUT,
        )
        return int((proc.stdout or "0").strip() or 0)
    except (FileNotFoundError, subprocess.TimeoutExpired, ValueError) as e:
        raise _SpotlightUnavailable(str(e)) from e


def _mdls_run(paths: list[str], names: list[str]) -> list[dict] | None:
    try:
        proc = subprocess.run(
            ["mdls", "-plist", "-", *names, *paths],
            capture_output=True, check=False, timeout=_SPOTLIGHT_TIMEOUT,
        )
    except (FileNotFoundError, subprocess.TimeoutExpired) as e:
        raise _SpotlightUnavailable(str(e)) from e
    if proc.returncode != 0 or not proc.stdout:
        return None
    try:
        data = plistlib.loads(proc.stdout)
    except Exception:
        return None
    return [data] if isinstance(data, dict) else list(data)


def _mdls(paths: list[str], attrs: Sequence[str]) -> list[dict]:
    """Read Spotlight attributes of many files in few mdls calls."""
    names: list[str] = []
    for a in attrs:
        names += ["-name", a]
    out: list[dict] = []
    for i in range(0, len(paths), 400):
        chunk = paths[i:i + 400]
        rows = _mdls_run(chunk, names)
        if rows is None:
            # A file vanished between mdfind and mdls (mail moved/deleted):
            # read one by one and skip what is gone.
            rows = []
            for p in chunk:
                r = _mdls_run([p], names)
                if r:
                    rows.extend(r)
        out.extend(rows)
    return out


def _first(val: Any) -> Any:
    if isinstance(val, (list, tuple)):
        return val[0] if val else None
    return val


def _category_names(val: Any, names: dict) -> list[str]:
    out = []
    for c in (val or []):
        n = names.get(str(c))
        if c and n and n not in out:
            out.append(n)
    return out


def _display_name(row: dict) -> str | None:
    n = row.get("kMDItemDisplayName") or row.get("kMDItemTitle")
    if not n or str(n).endswith((".olk15Message", ".olk15Event")):
        return None
    return n


_MAIL_ATTRS = [
    _MDP + "recordID", _MDP + "messageReceived", _MDP + "messageSent",
    _MDP + "unread", _MDP + "has_attachments", _MDP + "categories",
    _MDP + "author_email_addresses", "kMDItemAuthors", "kMDItemDisplayName",
]


def _list_emails_spotlight(
    folder: str | None, limit: int, unread_only: bool,
    from_filter: str | None, subject_filter: str | None,
    since: str | None, before: str | None, newest_first: bool,
) -> list[dict]:
    ctx = _run_jxa("spotlight_mail_context", {"folder": folder})
    fid = ctx.get("id")
    if fid is None:
        raise _SpotlightUnavailable("folder has no id")
    folder_q = f'kMDItemContentType == "{_MD_MESSAGE}" && {_MDP}folderID == {int(fid)}'
    base = folder_q + (f" && {_MDP}unread == 1" if unread_only else "")
    since_dt = _parse_iso(since) if since else None
    before_dt = _parse_iso(before) if before else None
    subj_q = (subject_filter or "").lower() or None
    from_q = (from_filter or "").lower() or None

    def matches(r: dict) -> bool:
        if subj_q and subj_q not in (_display_name(r) or "").lower():
            return False
        if from_q:
            who = " ".join(str(x) for x in (r.get("kMDItemAuthors") or [])
                           + (r.get(_MDP + "author_email_addresses") or []))
            if from_q not in who.lower():
                return False
        return True

    # Without `since`, look back in growing windows: the newest 50 of a
    # 36k inbox should not mean reading 36k index entries.
    if since_dt is not None or not newest_first:
        windows: list[_dt.datetime | None] = [since_dt]
    else:
        anchor = _to_utc(before_dt) if before_dt else _dt.datetime.now(_UTC)
        windows = [anchor - _dt.timedelta(days=d) for d in (7, 30, 120, 400)]
        windows.append(None)

    rows: list[dict] = []
    for lo in windows:
        q = base
        if lo is not None:
            q += f" && {_MDP}messageReceived >= {_md_time(lo)}"
        if before_dt is not None:
            q += f" && {_MDP}messageReceived < {_md_time(before_dt)}"
        rows = [r for r in _mdls(_mdfind(q), _MAIL_ATTRS) if matches(r)]
        if len(rows) >= limit:
            break

    if not rows and (ctx.get("count") or 0) > 0 and _mdfind_count(folder_q) == 0:
        # Outlook has mail there, Spotlight knows none: index missing.
        raise _SpotlightUnavailable("folder not in Spotlight index")

    epoch = _dt.datetime.min.replace(tzinfo=_UTC)
    rows.sort(key=lambda r: _md_utc(r.get(_MDP + "messageReceived")) or epoch,
              reverse=newest_first)
    return _mail_summaries(rows[:limit], ctx)


def _mail_summaries(rows: list[dict], ctx: dict) -> list[dict]:
    """Spotlight rows -> the dicts list_emails/search_emails return."""
    ids = [str(r.get(_MDP + "recordID")) for r in rows if r.get(_MDP + "recordID") is not None]
    extras = _run_jxa("email_extras", {"ids": ids}) if ids else {}
    cats = ctx.get("categories") or {}
    out = []
    for r in rows:
        rid = r.get(_MDP + "recordID")
        ex = extras.get(str(rid), {})
        out.append({
            "entry_id": str(rid) if rid is not None else None,
            "store_id": None,
            "subject": _display_name(r),
            "sender_name": _first(r.get("kMDItemAuthors")),
            "sender_email": _first(r.get(_MDP + "author_email_addresses")),
            "received_time": _md_iso(r.get(_MDP + "messageReceived")),
            "sent_time": _md_iso(r.get(_MDP + "messageSent")),
            "unread": bool(r.get(_MDP + "unread")),
            "has_attachments": bool(r.get(_MDP + "has_attachments")),
            "categories": _category_names(r.get(_MDP + "categories"), cats),
            "importance": ex.get("importance", "normal"),
            "folder": ctx.get("name"),
            "preview": ex.get("preview", ""),
        })
    return out


def _md_quote(text: str) -> str:
    """Escape user text for a quoted mdfind value with * wildcards around it."""
    out = []
    for ch in text:
        if ch in '\\"*?':
            out.append("\\" + ch)
        elif ch in "\r\n\t":
            out.append(" ")
        else:
            out.append(ch)
    return "".join(out)


def _search_query(folder_q: str, query: str) -> str:
    """Folder AND (text OR subject OR sender name OR sender address), like the
    JXA scan: case- and diacritics-insensitive substring match."""
    v = '"*' + _md_quote(query) + '*"cd'
    fields = ["kMDItemTextContent", "kMDItemDisplayName", "kMDItemAuthors",
              _MDP + "author_email_addresses"]
    return folder_q + " && (" + " || ".join(f"{f} == {v}" for f in fields) + ")"


def _search_emails_spotlight(query: str, folder: str | None, limit: int) -> list[dict]:
    ctx = _run_jxa("spotlight_mail_context", {"folder": folder})
    fid = ctx.get("id")
    if fid is None:
        raise _SpotlightUnavailable("folder has no id")
    folder_q = f'kMDItemContentType == "{_MD_MESSAGE}" && {_MDP}folderID == {int(fid)}'
    q = _search_query(folder_q, query)

    # Newest first, in growing windows: a common word can hit thousands of
    # mails, and mdls per hit is the expensive part.
    now = _dt.datetime.now(_UTC)
    windows: list[_dt.datetime | None] = [now - _dt.timedelta(days=d) for d in (30, 365)]
    windows.append(None)
    rows: list[dict] = []
    for lo in windows:
        wq = q if lo is None else q + f" && {_MDP}messageReceived >= {_md_time(lo)}"
        rows = _mdls(_mdfind(wq), _MAIL_ATTRS)
        if len(rows) >= limit:
            break

    if not rows and (ctx.get("count") or 0) > 0 and _mdfind_count(folder_q) == 0:
        raise _SpotlightUnavailable("folder not in Spotlight index")

    epoch = _dt.datetime.min.replace(tzinfo=_UTC)
    rows.sort(key=lambda r: _md_utc(r.get(_MDP + "messageReceived")) or epoch,
              reverse=True)
    return _mail_summaries(rows[:limit], ctx)


_EVENT_ATTRS = [
    _MDP + "recordID", _MDP + "startDate", _MDP + "endDate", _MDP + "allDay",
    _MDP + "recurring", _MDP + "recurrence", _MDP + "categories",
    _MDP + "parentRecordID",
    "kMDItemDisplayName", "kMDItemTitle", "kMDItemCoverage",
]
_UNTIL_UTC = re.compile(r"UNTIL=(\d{8}T\d{6})Z")


def _pick_calendars(cals: list[dict], name: str | None) -> list[int]:
    if name:
        ln = name.lower()
        ids = [c["id"] for c in cals if (c.get("name") or "").lower() == ln]
        if not ids:
            # Could be an account name — the JXA path knows how to map it.
            raise _SpotlightUnavailable("calendar name not found")
        return ids
    # The user's own calendar: an unnamed or "Calendar"/"Kalender" folder,
    # whichever holds the most events. Shared colleague calendars are
    # bigger and must not win.
    cands = [c for c in cals
             if (c.get("name") or "").lower() in ("", "calendar", "kalender")]
    best, best_n = None, 0
    for c in cands:
        n = _mdfind_count(
            f'kMDItemContentType == "{_MD_EVENT}" && {_MDP}folderID == {int(c["id"])}')
        if n > best_n:
            best, best_n = c["id"], n
    if best is None:
        raise _SpotlightUnavailable("no own calendar in Spotlight index")
    return [best]


def _expand_recurring(row: dict, ws: _dt.datetime, we: _dt.datetime,
                      skip_days: set | frozenset = frozenset(),
                      ) -> list[tuple[_dt.datetime, _dt.datetime]]:
    """Occurrences of a recurring master overlapping [ws, we) (UTC).

    ``skip_days`` holds local dates on which the series has an exception
    record of its own (a changed occurrence, ``parentRecordID`` = master);
    that record is listed instead. Occurrences deleted from a series leave
    no record in the index and cannot be filtered out here.
    """
    from dateutil.rrule import rrulestr  # local import: macOS-only dependency

    s = _md_utc(row.get(_MDP + "startDate"))
    e = _md_utc(row.get(_MDP + "endDate"))
    rule = row.get(_MDP + "recurrence")
    if s is None or e is None or not rule:
        return []
    dur = e - s
    # Recurrence rules run in local wall time (BYDAY, DST); UNTIL is either
    # floating local or UTC — normalize to floating local.
    s_loc = s.astimezone().replace(tzinfo=None)

    def _until_local(m: re.Match) -> str:
        u = _dt.datetime.strptime(m.group(1), "%Y%m%dT%H%M%S").replace(tzinfo=_UTC)
        return "UNTIL=" + u.astimezone().strftime("%Y%m%dT%H%M%S")

    rule = _UNTIL_UTC.sub(_until_local, rule)
    ws_loc = ws.astimezone().replace(tzinfo=None)
    we_loc = we.astimezone().replace(tzinfo=None)
    try:
        rr = rrulestr(rule, dtstart=s_loc)
        hits = rr.between(ws_loc - dur, we_loc, inc=True)
    except Exception:
        hits = [s_loc] if s < we and e > ws else []
    out = []
    for occ in hits:
        if occ.date() in skip_days:
            continue
        o_s = occ.astimezone(_UTC)  # naive local -> aware via system tz
        o_e = o_s + dur
        if o_s < we and o_e > ws:
            out.append((o_s, o_e))
    return out


def _list_calendar_spotlight(start: str, end: str, calendar: str | None,
                             limit: int) -> list[dict]:
    ws = _to_utc(_parse_iso(start))
    we = _to_utc(_parse_iso(end))
    ctx = _run_jxa("spotlight_calendar_context", {})
    fids = _pick_calendars(ctx.get("calendars") or [], calendar)
    cats = ctx.get("categories") or {}
    found: list[tuple[_dt.datetime, _dt.datetime, dict]] = []
    for fid in fids:
        base = f'kMDItemContentType == "{_MD_EVENT}" && {_MDP}folderID == {int(fid)}'
        single = (f"{base} && {_MDP}recurring == 0"
                  f" && {_MDP}startDate < {_md_time(we)}"
                  f" && {_MDP}endDate > {_md_time(ws)}")
        exceptions: dict[int, set] = {}
        for r in _mdls(_mdfind(single), _EVENT_ATTRS):
            s = _md_utc(r.get(_MDP + "startDate"))
            e = _md_utc(r.get(_MDP + "endDate"))
            if s is None or e is None:
                continue
            found.append((s, e, r))
            parent = r.get(_MDP + "parentRecordID")
            if parent is not None:
                exceptions.setdefault(parent, set()).add(s.astimezone().date())
        series = f"{base} && {_MDP}recurring == 1 && {_MDP}startDate < {_md_time(we)}"
        for r in _mdls(_mdfind(series), _EVENT_ATTRS):
            skip = exceptions.get(r.get(_MDP + "recordID"), set())
            for s, e in _expand_recurring(r, ws, we, skip):
                found.append((s, e, r))
    found.sort(key=lambda t: t[0])
    found = found[:limit]

    ids = sorted({str(r.get(_MDP + "recordID")) for _, _, r in found
                  if r.get(_MDP + "recordID") is not None})
    extras = _run_jxa("event_extras", {"ids": ids}) if ids else {}
    out = []
    for s, e, r in found:
        rid = r.get(_MDP + "recordID")
        ex = extras.get(str(rid), {})
        out.append({
            "entry_id": str(rid) if rid is not None else None,
            "store_id": None,
            "subject": _display_name(r),
            "location": ex.get("location") or r.get("kMDItemCoverage"),
            "start": _md_iso(s),
            "end": _md_iso(e),
            "all_day": bool(r.get(_MDP + "allDay")),
            "body": ex.get("body", ""),
            "categories": _category_names(r.get(_MDP + "categories"), cats),
            "organizer": ex.get("organizer"),
            "is_recurring": bool(r.get(_MDP + "recurring")),
        })
    return out


# ---- public API ----------------------------------------------------------

def list_accounts() -> list[dict]:
    return _run_jxa("list_accounts")


def list_folders(account: str | None = None, recursive: bool = True) -> list[dict]:
    return _run_jxa("list_folders", {"account": account, "recursive": recursive})


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
    if since:
        _parse_iso(since)
    if before:
        _parse_iso(before)
    if _spotlight_enabled():
        try:
            return _list_emails_spotlight(
                _canonical_folder_path(folder), limit, unread_only,
                from_filter, subject_filter, since, before, newest_first)
        except _SpotlightUnavailable:
            pass
    return _run_jxa("list_emails", {
        "folder": _canonical_folder_path(folder),
        "limit": limit,
        "unread_only": unread_only,
        "from_filter": from_filter,
        "subject_filter": subject_filter,
        "since": since,
        "before": before,
        "newest_first": newest_first,
    })


def read_email(entry_id: str, store_id: str | None = None) -> dict:
    return _run_jxa("read_email", {"entry_id": entry_id})


def search_emails(query: str, folder: str | None = None, limit: int = 50) -> list[dict]:
    if query and _spotlight_enabled():
        try:
            return _search_emails_spotlight(query, _canonical_folder_path(folder), limit)
        except _SpotlightUnavailable:
            pass
    return _run_jxa("search_emails", {
        "query": query,
        "folder": _canonical_folder_path(folder),
        "limit": limit,
    })


def move_email(entry_id: str, target_folder: str, store_id: str | None = None) -> dict:
    return _run_jxa("move_email", {
        "entry_id": entry_id,
        "target_folder": _canonical_folder_path(target_folder),
    })


def delete_email(entry_id: str, store_id: str | None = None, permanent: bool = False) -> dict:
    return _run_jxa("delete_email", {"entry_id": entry_id, "permanent": permanent})


def _validate_attachments(paths: Sequence[str] | None) -> None:
    if paths:
        for p in paths:
            _validate_attachment_path(p)


def create_draft(
    to: list[str] | None = None,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    subject: str | None = None,
    body: str | None = None,
    html_body: str | None = None,
    attachments: list[str] | None = None,
    reply_to_entry_id: str | None = None,
    reply_to_store_id: str | None = None,
    reply_mode: Literal["reply", "reply_all", "forward"] | None = None,
    importance: Literal["low", "normal", "high"] = "normal",
    categories: list[str] | None = None,
) -> dict:
    _validate_attachments(attachments)
    if html_body:
        html_body = _sanitize_html_body(html_body)
    msg_id = _create_draft_as(
        subject=subject, body=body, html_body=html_body,
        importance=importance,
        to=to, cc=cc, bcc=bcc,
        reply_to_entry_id=reply_to_entry_id, reply_mode=reply_mode,
        categories=categories, attachments=attachments,
    )
    return _summarize_message_as(msg_id)


def update_draft(
    entry_id: str,
    store_id: str | None = None,
    to: list[str] | None = None,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    subject: str | None = None,
    body: str | None = None,
    html_body: str | None = None,
    add_attachments: list[str] | None = None,
    importance: Literal["low", "normal", "high"] | None = None,
    categories: list[str] | None = None,
) -> dict:
    _validate_attachments(add_attachments)
    if html_body:
        html_body = _sanitize_html_body(html_body)
    _update_draft_as(
        entry_id,
        subject=subject, body=body, html_body=html_body,
        importance=importance,
        to=to, cc=cc, bcc=bcc,
        add_attachments=add_attachments,
        categories=categories,
    )
    return _summarize_message_as(entry_id)


def send_email(
    entry_id: str | None = None,
    store_id: str | None = None,
    to: list[str] | None = None,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    subject: str | None = None,
    body: str | None = None,
    html_body: str | None = None,
    attachments: list[str] | None = None,
    importance: Literal["low", "normal", "high"] = "normal",
) -> dict:
    _validate_attachments(attachments)
    if html_body:
        html_body = _sanitize_html_body(html_body)
    return _run_jxa("send_email", {
        "entry_id": entry_id,
        "to": to, "cc": cc, "bcc": bcc,
        "subject": subject, "body": body, "html_body": html_body,
        "attachments": attachments,
        "importance": importance,
    })


def list_categories() -> list[dict]:
    return _run_jxa("list_categories")


def set_email_categories(
    entry_id: str,
    categories: list[str],
    mode: Literal["replace", "add", "remove"] = "replace",
    store_id: str | None = None,
) -> dict:
    return _set_categories_as(entry_id, categories, mode, "message")


def set_event_categories(
    entry_id: str,
    categories: list[str],
    mode: Literal["replace", "add", "remove"] = "replace",
    store_id: str | None = None,
) -> dict:
    return _set_categories_as(entry_id, categories, mode, "event")


def list_calendar_events(
    start: str,
    end: str,
    calendar: str | None = None,
    limit: int = 100,
) -> list[dict]:
    _parse_iso(start); _parse_iso(end)
    if _spotlight_enabled():
        try:
            return _list_calendar_spotlight(start, end, calendar, limit)
        except _SpotlightUnavailable:
            pass
    return _run_jxa("list_calendar_events", {
        "start": start, "end": end, "calendar": calendar, "limit": limit,
    })


def create_calendar_event(
    subject: str,
    start: str,
    end: str,
    body: str | None = None,
    location: str | None = None,
    attendees: list[str] | None = None,
    categories: list[str] | None = None,
    all_day: bool = False,
    is_meeting: bool = False,
    reminder_minutes: int | None = None,
) -> dict:
    _parse_iso(start); _parse_iso(end)
    summary = _run_jxa("create_calendar_event", {
        "subject": subject, "start": start, "end": end,
        "body": body, "location": location,
        "all_day": all_day, "is_meeting": is_meeting,
        "reminder_minutes": reminder_minutes,
    })
    event_id = summary["entry_id"]
    if categories:
        cat_result = _set_categories_as(event_id, categories, "replace", "event")
        summary["categories"] = cat_result["categories"]
    if attendees:
        for addr in attendees:
            _add_attendee(event_id, addr, optional=False)
        summary["attendees_added"] = attendees
    return summary


def update_calendar_event(
    entry_id: str,
    store_id: str | None = None,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    body: str | None = None,
    location: str | None = None,
    add_attendees: list[str] | None = None,
    categories: list[str] | None = None,
    send_update: bool = False,
) -> dict:
    if start: _parse_iso(start)
    if end: _parse_iso(end)
    summary = _run_jxa("update_calendar_event", {
        "entry_id": entry_id, "subject": subject,
        "start": start, "end": end,
        "body": body, "location": location,
        "send_update": send_update,
    })
    if categories is not None:
        cat_result = _set_categories_as(entry_id, categories, "replace", "event")
        summary["categories"] = cat_result["categories"]
    if add_attendees:
        for addr in add_attendees:
            _add_attendee(entry_id, addr, optional=False)
        summary["attendees_added"] = add_attendees
    return summary


def delete_calendar_event(entry_id: str, store_id: str | None = None) -> dict:
    return _run_jxa("delete_calendar_event", {"entry_id": entry_id})


def respond_to_invite(
    entry_id: str,
    response: Literal["accept", "tentative", "decline"],
    send_response: bool = True,
    store_id: str | None = None,
) -> dict:
    return _run_jxa("respond_to_invite", {
        "entry_id": entry_id, "response": response, "send_response": send_response,
    })
