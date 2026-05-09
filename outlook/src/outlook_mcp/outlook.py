"""Platform router for the Outlook backend.

Re-exports the public API from the platform-specific implementation:

- Windows  → ``outlook_win`` (COM via pywin32)
- macOS    → ``outlook_mac`` (AppleScript / JXA via osascript)
"""
from __future__ import annotations

import sys

_PUBLIC_NAMES = (
    "OutlookError",
    "list_accounts",
    "list_folders",
    "list_emails",
    "read_email",
    "search_emails",
    "move_email",
    "delete_email",
    "create_draft",
    "update_draft",
    "send_email",
    "list_categories",
    "set_email_categories",
    "set_event_categories",
    "list_calendar_events",
    "create_calendar_event",
    "update_calendar_event",
    "delete_calendar_event",
    "respond_to_invite",
)

if sys.platform == "win32":
    from . import outlook_win as _impl
elif sys.platform == "darwin":
    from . import outlook_mac as _impl
else:
    raise RuntimeError(
        f"outlook_mcp: unsupported platform {sys.platform!r}. "
        "Supported platforms: win32 (COM), darwin (AppleScript)."
    )

# Re-export the public surface so ``from . import outlook as ol`` works
# identically to the legacy single-file backend.
for _name in _PUBLIC_NAMES:
    globals()[_name] = getattr(_impl, _name)

__all__ = list(_PUBLIC_NAMES)
