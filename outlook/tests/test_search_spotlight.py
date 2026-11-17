"""search_emails on macOS: Spotlight fast path and its JXA fallback.

mdfind/mdls and the JXA bridge are replaced by fakes, so this runs without
Outlook. What it pins:
  - user text is escaped for the mdfind value (quotes, backslash, wildcards)
  - the query ORs text, subject, sender name and sender address in the folder
  - results come newest first, cut to `limit`, windows widen only when needed
  - Spotlight trouble falls back to the JXA scan
"""
from __future__ import annotations

import datetime as dt

import pytest

from outlook_mcp import outlook_mac as om

P = om._MDP


def _row(rid: int, received: dt.datetime, subject: str = "s") -> dict:
    return {
        P + "recordID": rid,
        P + "messageReceived": received,
        P + "messageSent": received,
        P + "unread": 0,
        P + "has_attachments": 0,
        P + "categories": [],
        P + "author_email_addresses": ["a@b.de"],
        "kMDItemAuthors": ["A B"],
        "kMDItemDisplayName": subject,
    }


@pytest.fixture
def fake(monkeypatch):
    state = {"queries": [], "rows": [], "jxa": []}

    def run_jxa(action, payload):
        state["jxa"].append(action)
        if action == "spotlight_mail_context":
            return {"id": 116, "name": "Posteingang", "count": 10, "categories": {}}
        if action == "email_extras":
            return {i: {"importance": "high" if i == "2" else "normal", "preview": "p" + i}
                    for i in payload["ids"]}
        if action == "search_emails":
            return [{"entry_id": "jxa"}]
        raise AssertionError(action)

    def mdfind(q):
        state["queries"].append(q)
        rows = state["rows"]
        if "messageReceived >=" in q:
            lo = q.split("$time.iso(")[1].split(")")[0]
            lo = dt.datetime.strptime(lo, "%Y-%m-%dT%H:%M:%SZ")
            rows = [r for r in rows if r[P + "messageReceived"] >= lo]
        return [str(r[P + "recordID"]) for r in rows]

    def mdls(paths, attrs):
        by_id = {str(r[P + "recordID"]): r for r in state["rows"]}
        return [by_id[p] for p in paths]

    monkeypatch.setattr(om, "_run_jxa", run_jxa)
    monkeypatch.setattr(om, "_mdfind", mdfind)
    monkeypatch.setattr(om, "_mdls", mdls)
    monkeypatch.setattr(om, "_mdfind_count", lambda q: 10)
    monkeypatch.setattr(om, "_spotlight_enabled", lambda: True)
    return state


def test_quote_escapes_mdfind_specials():
    assert om._md_quote('a"b\\c*d?e') == 'a\\"b\\\\c\\*d\\?e'
    assert om._md_quote("zwei\nzeilen") == "zwei zeilen"


def test_query_ors_all_fields_inside_folder():
    q = om._search_query("FOLDER", "Einladung")
    assert q.startswith("FOLDER && (")
    for field in ("kMDItemTextContent", "kMDItemDisplayName", "kMDItemAuthors",
                  P + "author_email_addresses"):
        assert f'{field} == "*Einladung*"cd' in q
    assert q.count(" || ") == 3


def test_newest_first_and_limit(fake):
    now = dt.datetime.now(dt.timezone.utc).replace(tzinfo=None)
    fake["rows"] = [_row(i, now - dt.timedelta(days=i)) for i in range(1, 6)]
    out = om.search_emails("x", limit=3)
    assert [m["entry_id"] for m in out] == ["1", "2", "3"]
    assert out[1]["importance"] == "high"
    assert out[0]["preview"] == "p1"
    assert out[0]["folder"] == "Posteingang"
    assert "search_emails" not in fake["jxa"]


def test_window_widens_until_limit(fake):
    now = dt.datetime.now(dt.timezone.utc).replace(tzinfo=None)
    # one hit in the last 30 days, two older than a year
    fake["rows"] = [_row(1, now - dt.timedelta(days=3)),
                    _row(2, now - dt.timedelta(days=500)),
                    _row(3, now - dt.timedelta(days=600))]
    out = om.search_emails("x", limit=2)
    assert [m["entry_id"] for m in out] == ["1", "2"]
    assert len(fake["queries"]) == 3          # 30 d, 365 d, unbounded
    assert "messageReceived" not in fake["queries"][-1]


def test_stops_at_first_window_that_suffices(fake):
    now = dt.datetime.now(dt.timezone.utc).replace(tzinfo=None)
    fake["rows"] = [_row(i, now - dt.timedelta(days=i)) for i in range(1, 4)]
    om.search_emails("x", limit=2)
    assert len(fake["queries"]) == 1


def test_spotlight_failure_falls_back_to_jxa(fake, monkeypatch):
    def broken(q):
        raise om._SpotlightUnavailable("index off")
    monkeypatch.setattr(om, "_mdfind", broken)
    assert om.search_emails("x") == [{"entry_id": "jxa"}]


def test_unindexed_folder_falls_back_to_jxa(fake, monkeypatch):
    fake["rows"] = []
    monkeypatch.setattr(om, "_mdfind_count", lambda q: 0)
    assert om.search_emails("x") == [{"entry_id": "jxa"}]


def test_empty_query_skips_spotlight(fake):
    om.search_emails("")
    assert fake["jxa"] == ["search_emails"]
