"""Tests for German locale support in the Outlook COM wrapper.

These tests cover the pure-logic parts (no COM/Outlook needed):
  - _format_filter_dt: locale-independent date formatting
  - _NAMED_FOLDERS: German alias resolution
  - _walk_subfolders: case-insensitive fallback
  - _parse_categories: semicolon delimiter (common in German locale)
"""
from __future__ import annotations

import datetime as dt
from unittest.mock import MagicMock

import pytest

# The module imports pywin32 at the top level, which isn't available in CI.
# Patch the win32 modules before importing outlook.
import sys
sys.modules.setdefault("pythoncom", MagicMock())
sys.modules.setdefault("pywintypes", MagicMock())
sys.modules.setdefault("win32com", MagicMock())
sys.modules.setdefault("win32com.client", MagicMock())

from outlook_mcp.outlook import (
    _format_filter_dt,
    _escape_dasl,
    _validate_attachment_path,
    _sanitize_html_body,
    _NAMED_FOLDERS,
    _parse_categories,
    _walk_subfolders,
    OL_FOLDER_INBOX,
    OL_FOLDER_SENT,
    OL_FOLDER_DRAFTS,
    OL_FOLDER_DELETED,
    OL_FOLDER_JUNK,
    OL_FOLDER_OUTBOX,
    OL_FOLDER_CALENDAR,
    OL_FOLDER_CONTACTS,
    OL_FOLDER_TASKS,
    OutlookError,
)


# ---------- _format_filter_dt ----------

class TestFormatFilterDt:
    def test_basic_datetime_pm(self):
        d = dt.datetime(2026, 5, 6, 14, 30, 0)
        assert _format_filter_dt(d) == "5/6/2026 2:30 PM"

    def test_midnight(self):
        d = dt.datetime(2026, 1, 1, 0, 0, 0)
        assert _format_filter_dt(d) == "1/1/2026 12:00 AM"

    def test_morning(self):
        d = dt.datetime(2026, 3, 7, 9, 5, 0)
        assert _format_filter_dt(d) == "3/7/2026 9:05 AM"

    def test_noon(self):
        d = dt.datetime(2026, 5, 6, 12, 0, 0)
        assert _format_filter_dt(d) == "5/6/2026 12:00 PM"

    def test_end_of_day(self):
        d = dt.datetime(2026, 12, 31, 23, 59, 0)
        assert _format_filter_dt(d) == "12/31/2026 11:59 PM"

    def test_no_strftime_p_used(self):
        """Ensure AM/PM is hardcoded, not from locale-sensitive strftime(%p)."""
        d = dt.datetime(2026, 5, 6, 15, 0, 0)
        result = _format_filter_dt(d)
        assert result.endswith("PM")
        d2 = dt.datetime(2026, 5, 6, 3, 0, 0)
        result2 = _format_filter_dt(d2)
        assert result2.endswith("AM")


# ---------- _NAMED_FOLDERS: German aliases ----------

class TestNamedFoldersGerman:
    @pytest.mark.parametrize("alias,expected", [
        ("posteingang", OL_FOLDER_INBOX),
        ("gesendete elemente", OL_FOLDER_SENT),
        ("gesendet", OL_FOLDER_SENT),
        ("entwürfe", OL_FOLDER_DRAFTS),
        ("gelöschte elemente", OL_FOLDER_DELETED),
        ("papierkorb", OL_FOLDER_DELETED),
        ("junk-e-mail", OL_FOLDER_JUNK),
        ("postausgang", OL_FOLDER_OUTBOX),
        ("kalender", OL_FOLDER_CALENDAR),
        ("kontakte", OL_FOLDER_CONTACTS),
        ("aufgaben", OL_FOLDER_TASKS),
    ])
    def test_german_alias_resolves(self, alias, expected):
        assert _NAMED_FOLDERS[alias] == expected

    def test_english_aliases_still_work(self):
        assert _NAMED_FOLDERS["inbox"] == OL_FOLDER_INBOX
        assert _NAMED_FOLDERS["sent"] == OL_FOLDER_SENT
        assert _NAMED_FOLDERS["drafts"] == OL_FOLDER_DRAFTS
        assert _NAMED_FOLDERS["deleted"] == OL_FOLDER_DELETED
        assert _NAMED_FOLDERS["junk"] == OL_FOLDER_JUNK

    def test_german_and_english_map_to_same_constant(self):
        assert _NAMED_FOLDERS["inbox"] == _NAMED_FOLDERS["posteingang"]
        assert _NAMED_FOLDERS["sent"] == _NAMED_FOLDERS["gesendet"]
        assert _NAMED_FOLDERS["drafts"] == _NAMED_FOLDERS["entwürfe"]
        assert _NAMED_FOLDERS["calendar"] == _NAMED_FOLDERS["kalender"]


# ---------- _walk_subfolders ----------

def _make_folder(name: str, children: list | None = None):
    """Create a mock COM folder with .Name and .Folders."""
    folder = MagicMock()
    folder.Name = name
    child_list = children or []
    folder.Folders.Count = len(child_list)
    folder.Folders.Item = lambda i: child_list[i - 1] if isinstance(i, int) else _item_by_name(child_list, i)
    return folder


def _item_by_name(children, name):
    for c in children:
        if c.Name == name:
            return c
    import pywintypes
    raise pywintypes.com_error("not found", None, None, None)


class TestWalkSubfolders:
    def test_exact_name_match(self):
        child = _make_folder("Projects")
        parent = _make_folder("Inbox", [child])
        result = _walk_subfolders(parent, ["Projects"])
        assert result.Name == "Projects"

    def test_case_insensitive_fallback(self):
        # Simulate German Outlook: folder is "Posteingang" but caller passes "posteingang"
        child = _make_folder("Posteingang")
        parent = _make_folder("Root", [child])
        # Exact match will fail (COM is case-sensitive for Item(str)),
        # but our fallback should find it case-insensitively
        parent.Folders.Item = lambda i: (
            child if isinstance(i, int) and i == 1
            else (_ for _ in ()).throw(type(sys.modules["pywintypes"]).com_error("not found", None, None, None))
        )
        # Re-mock to simulate COM behavior: Item(str) raises, Item(int) works
        import pywintypes
        com_error = type("com_error", (Exception,), {})
        pywintypes.com_error = com_error

        def item_access(i):
            if isinstance(i, int):
                return child
            raise com_error("not found")

        parent.Folders.Item = item_access
        parent.Folders.Count = 1

        result = _walk_subfolders(parent, ["posteingang"])
        assert result.Name == "Posteingang"

    def test_nested_subfolders(self):
        grandchild = _make_folder("Q1")
        child = _make_folder("Projects", [grandchild])
        parent = _make_folder("Inbox", [child])
        result = _walk_subfolders(parent, ["Projects", "Q1"])
        assert result.Name == "Q1"

    def test_not_found_raises_outlook_error(self):
        import pywintypes
        com_error = type("com_error", (Exception,), {})
        pywintypes.com_error = com_error

        parent = _make_folder("Inbox", [])
        parent.Folders.Item = lambda i: (_ for _ in ()).throw(com_error("nope"))
        parent.Folders.Count = 0

        with pytest.raises(OutlookError, match="not found"):
            _walk_subfolders(parent, ["NonExistent"])


# ---------- _parse_categories (semicolon support) ----------

# ---------- _escape_dasl ----------

class TestEscapeDasl:
    def test_single_quotes_doubled(self):
        assert _escape_dasl("O'Brien") == "O''Brien"

    def test_percent_escaped(self):
        assert _escape_dasl("100%") == "100[%]"

    def test_underscore_escaped(self):
        assert _escape_dasl("file_name") == "file[_]name"

    def test_combined(self):
        assert _escape_dasl("it's 50% done_now") == "it''s 50[%] done[_]now"

    def test_plain_text_unchanged(self):
        assert _escape_dasl("hello world") == "hello world"

    def test_injection_attempt_neutralized(self):
        malicious = "%' OR 1=1 OR '"
        escaped = _escape_dasl(malicious)
        assert "OR" in escaped  # text preserved
        assert escaped.startswith("[%]")  # but % is escaped


# ---------- _validate_attachment_path ----------

class TestValidateAttachmentPath:
    @pytest.mark.parametrize("path", [
        r"C:\Users\me\.ssh\id_rsa",
        r"C:\Users\me\.ssh\id_ed25519",
        r"C:\Users\me\.aws\credentials",
        r"C:\Users\me\.gnupg\private.key",
        r"C:\Users\me\.kube\config",
        r"C:\project\server.pem",
        r"C:\project\private.key",
        r"C:\project\cert.pfx",
        r"C:\project\keystore.p12",
        r"C:\project\credentials.json",
        r"C:\project\.env",
        r"C:\project\.env.production",
        r"C:\Users\me\.netrc",
    ])
    def test_blocks_sensitive_paths(self, path):
        with pytest.raises(OutlookError, match="sensitive"):
            _validate_attachment_path(path)

    @pytest.mark.parametrize("path", [
        r"C:\Users\me\Documents\report.pdf",
        r"C:\Users\me\Desktop\photo.jpg",
        r"C:\project\readme.md",
        r"C:\project\data.csv",
    ])
    def test_allows_normal_paths(self, path):
        _validate_attachment_path(path)  # should not raise


# ---------- _sanitize_html_body ----------

class TestSanitizeHtmlBody:
    def test_strips_script_tag(self):
        html = '<p>Hello</p><script>alert("xss")</script>'
        result = _sanitize_html_body(html)
        assert "<script" not in result
        assert "<p>Hello</p>" in result

    def test_strips_iframe(self):
        html = '<iframe src="https://evil.com"></iframe>'
        assert "<iframe" not in _sanitize_html_body(html)

    def test_strips_object_embed_applet(self):
        for tag in ["object", "embed", "applet"]:
            html = f'<{tag} data="x"></{tag}>'
            assert f"<{tag}" not in _sanitize_html_body(html)

    def test_strips_form(self):
        html = '<form action="https://evil.com"><input></form>'
        assert "<form" not in _sanitize_html_body(html)

    def test_strips_meta_http_equiv(self):
        html = '<meta http-equiv="refresh" content="0;url=evil.com">'
        assert "<meta" not in _sanitize_html_body(html)

    def test_preserves_safe_html(self):
        html = '<div><p>Hello <b>World</b></p><img src="logo.png"></div>'
        assert _sanitize_html_body(html) == html

    def test_case_insensitive(self):
        html = '<SCRIPT>alert("xss")</SCRIPT>'
        assert "<SCRIPT" not in _sanitize_html_body(html)


# ---------- _parse_categories ----------

class TestParseCategories:
    def test_comma_separated(self):
        assert _parse_categories("Red,Blue,Green") == ["Red", "Blue", "Green"]

    def test_semicolon_separated(self):
        assert _parse_categories("Rot;Blau;Grün") == ["Rot", "Blau", "Grün"]

    def test_mixed_delimiters(self):
        assert _parse_categories("A,B;C") == ["A", "B", "C"]

    def test_empty_string(self):
        assert _parse_categories("") == []

    def test_none(self):
        assert _parse_categories(None) == []

    def test_whitespace_trimmed(self):
        assert _parse_categories(" Red , Blue ; Green ") == ["Red", "Blue", "Green"]
