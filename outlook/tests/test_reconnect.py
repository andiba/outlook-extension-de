"""Tests for the auto-reconnect path in the Windows COM backend.

Covers the logic that recovers from a dead/stale COM reference after Outlook
is restarted (crash, update, or manual close/open):
  - _is_stale_com_error: HRESULT classification
  - _with_reconnect: retry-once-on-stale decorator
  - _app_ns: cheap health check + transparent reconnect

COM itself is hard to mock, so we patch ``pywintypes.com_error`` to a real
exception class and drive the pure control flow with fakes.
"""
from __future__ import annotations

from unittest.mock import MagicMock

import pytest

# pywin32 isn't importable in CI — stub the modules before importing the backend.
import sys
sys.modules.setdefault("pythoncom", MagicMock())
sys.modules.setdefault("pywintypes", MagicMock())
sys.modules.setdefault("win32com", MagicMock())
sys.modules.setdefault("win32com.client", MagicMock())

import pywintypes  # noqa: E402  (the stub, or the real thing on Windows)


class _ComError(Exception):
    """Stand-in for pywintypes.com_error: args == (hresult, strerror, excepinfo, argerror)."""


# The decorator/helpers use ``except pywintypes.com_error`` at call time, so
# swapping this in is enough to make them catch our fake errors.
pywintypes.com_error = _ComError

from outlook_mcp import outlook_win as ow  # noqa: E402

# HRESULTs the backend treats as "connection is dead, reconnect".
STALE = -2147023174        # RPC_S_SERVER_UNAVAILABLE
STALE_DISCONNECTED = -2147417848
STALE_OBJNOTCONNECTED = -2147220995
STALE_SERVER_EXEC = -2146959355
NOT_STALE = -2147221233    # MAPI_E_NOT_FOUND-ish — a genuine error, must propagate


def _com(hresult, *, excepinfo=None):
    return _ComError(hresult, "boom", excepinfo, None)


@pytest.fixture(autouse=True)
def _reset_module_state():
    """Isolate the module-global connection cache between tests.

    Also (re)install our fake com_error each time: other test modules mutate
    the shared ``pywintypes.com_error`` attribute, so pin it right before use.
    """
    pywintypes.com_error = _ComError
    ow._app = None
    ow._ns = None
    ow._com_initialized = True  # pretend CoInitialize already happened
    yield
    ow._app = None
    ow._ns = None


# ---------- _is_stale_com_error ----------

class TestIsStaleComError:
    @pytest.mark.parametrize("hr", [
        STALE, STALE_DISCONNECTED, STALE_OBJNOTCONNECTED, STALE_SERVER_EXEC,
    ])
    def test_stale_hresults_detected(self, hr):
        assert ow._is_stale_com_error(_com(hr)) is True

    def test_hresult_hidden_in_excepinfo(self):
        # EXCEPINFO tuple; scode sits at index 5.
        excepinfo = (0, "src", "desc", "help", 0, STALE)
        assert ow._is_stale_com_error(_com(0, excepinfo=excepinfo)) is True

    def test_unrelated_hresult_is_not_stale(self):
        assert ow._is_stale_com_error(_com(NOT_STALE)) is False

    def test_empty_args_is_not_stale(self):
        assert ow._is_stale_com_error(_ComError()) is False


# ---------- _with_reconnect ----------

class TestWithReconnect:
    def test_success_first_try_no_retry(self):
        calls = []

        @ow._with_reconnect
        def fn():
            calls.append(1)
            return "ok"

        assert fn() == "ok"
        assert len(calls) == 1

    def test_stale_then_success_retries_once_and_resets(self):
        ow._app = object()  # sentinel live connection
        ow._ns = object()
        calls = []

        @ow._with_reconnect
        def fn():
            calls.append(1)
            if len(calls) == 1:
                raise _com(STALE)
            return "recovered"

        assert fn() == "recovered"
        assert len(calls) == 2
        # Stale handler must have dropped the cached refs before the retry.
        assert ow._app is None
        assert ow._ns is None

    def test_stale_twice_raises_outlook_error(self):
        @ow._with_reconnect
        def fn():
            raise _com(STALE_DISCONNECTED)

        with pytest.raises(ow.OutlookError, match="Reconnect fehlgeschlagen"):
            fn()

    def test_non_stale_com_error_propagates_without_retry(self):
        calls = []

        @ow._with_reconnect
        def fn():
            calls.append(1)
            raise _com(NOT_STALE)

        with pytest.raises(_ComError):
            fn()
        assert len(calls) == 1  # no retry for genuine errors


# ---------- _app_ns health check ----------

class _LiveApp:
    Version = "16.0.0"


class _DeadApp:
    """Cached proxy pointing at a dead Outlook — any property access raises."""

    def __init__(self, hresult=STALE):
        self._hr = hresult

    @property
    def Version(self):
        raise _com(self._hr)


class TestAppNsHealthCheck:
    def test_cold_start_dispatches(self, monkeypatch):
        connected = []

        def fake_connect():
            ow._app = _LiveApp()
            ow._ns = "ns"
            connected.append(1)

        monkeypatch.setattr(ow, "_connect", fake_connect)
        app, ns = ow._app_ns()
        assert isinstance(app, _LiveApp)
        assert connected == [1]

    def test_healthy_connection_is_reused(self, monkeypatch):
        live = _LiveApp()
        ow._app = live
        ow._ns = "ns"
        monkeypatch.setattr(
            ow, "_connect",
            lambda: pytest.fail("should not reconnect a healthy connection"),
        )
        app, _ = ow._app_ns()
        assert app is live

    def test_stale_connection_reconnects(self, monkeypatch):
        fresh = _LiveApp()

        def fake_connect():
            ow._app = fresh
            ow._ns = "fresh-ns"

        monkeypatch.setattr(ow, "_connect", fake_connect)
        ow._app = _DeadApp()  # health check will raise a stale error
        ow._ns = "old-ns"

        app, ns = ow._app_ns()
        assert app is fresh
        assert ns == "fresh-ns"

    def test_non_stale_health_error_propagates(self, monkeypatch):
        monkeypatch.setattr(
            ow, "_connect",
            lambda: pytest.fail("must not reconnect on a genuine error"),
        )
        ow._app = _DeadApp(hresult=NOT_STALE)
        ow._ns = "ns"

        with pytest.raises(_ComError):
            ow._app_ns()
