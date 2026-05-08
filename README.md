# Claude Local Connectors — Outlook + Obsidian

Shareable Claude Code / Cowork extensions that connect Claude to **local desktop applications** — no cloud APIs, no OAuth, no external services.

**Windows only.**

| Extension | What it does | Requirements |
|---|---|---|
| **[Outlook](outlook/)** | Email, calendar, categories via COM | Outlook desktop + [uv](https://docs.astral.sh/uv/) |
| **[Obsidian](obsidian-dxt/)** | Read, search, create, edit vault notes via MCP Tools | Obsidian + [MCP Tools plugin](https://github.com/jacksteamdev/obsidian-mcp-tools) |

---

## Outlook

Lets Claude drive the **locally-installed Microsoft Outlook desktop client** via COM. No Azure AD registration, no OAuth, no cloud. Works against whatever accounts your Outlook profile already has.

## What it does

| Surface | Capabilities |
|---|---|
| **Email** | list, read, search, move, delete, draft, reply, reply-all, forward, update draft, send |
| **Calendar** | list events (with recurrence expansion), create, update, delete, accept/decline invites |
| **Categories** | list master list, apply to emails or events (replace/add/remove) |
| **Accounts & folders** | enumerate accounts and full folder trees |

Plus:

- **`outlook-assistant`** skill — auto-loads when the user mentions email, calendar, or meetings, teaching Claude the safe-default behaviors (draft before send, soft delete, confirm batch ops).
- **Slash commands:** `/triage-inbox`, `/daily-digest`, `/draft-reply`, `/clean-inbox`.

## Install

### Prerequisites

- **Windows** with Outlook desktop installed and a profile configured
- [**uv**](https://docs.astral.sh/uv/getting-started/installation/) on PATH (handles Python + dependencies)

### Option A — Claude Code (full plugin: MCP tools + skill + slash commands)

```powershell
claude plugin marketplace add andiba/outlook-extension-de
claude plugin install outlook@outlook-tools
```

Restart Claude Code. Run `/mcp` to confirm the `outlook` server shows up and `/plugin` to see the plugin listed. On first run, `uv` creates a venv inside the plugin directory and installs dependencies; subsequent launches reuse it.

This option gives you the full experience: the MCP server, the `outlook-assistant` skill, and slash commands (`/triage-inbox`, `/daily-digest`, `/draft-reply`, `/clean-inbox`).

### Option B — Cowork / Claude Desktop (MCP server only, via `.mcpb` bundle)

Download the latest `outlook-<version>.mcpb` from the [Releases](https://github.com/andiba/outlook-extension-de/releases) page, then open Claude Desktop: **Settings → Extensions → Install from file** and pick the `.mcpb`.

> The desktop bundle contains the MCP server only. Skills and slash commands are a Claude Code feature and are not loaded in the desktop app.

### Both Claude Code + Cowork

Claude Code and Cowork use **separate plugin registries**. If you want the extension in both, you need to install it twice:

1. **Claude Code** — Option A above
2. **Cowork** — Option B above (`.mcpb` from the Releases page)

There is no single command that covers both.

## Use it

```
You: /daily-digest
Claude: [calls list_calendar_events + list_emails, produces brief]

You: Draft a reply to Sarah saying I'll get to it Friday
Claude: [search_emails → read_email → create_draft, surfaces draft]

You: Looks good, send it
Claude: [send_email]
```

Or just talk to it: *"Move all unread newsletters from this week to the Archive folder"* — the `outlook-assistant` skill teaches Claude to confirm batch operations before executing.

## Develop locally

The plugin lives under [outlook/](outlook/). To iterate:

```powershell
cd outlook
uv sync                   # creates .venv, installs deps
uv run outlook-mcp        # runs the MCP server over stdio (testing)
```

To load your local copy into Claude Code without publishing:

```powershell
# Add this directory as a local marketplace
claude plugin marketplace add "C:\Users\gabri\Outlook Extension"
claude plugin install outlook@outlook-tools
```

### Layout

```
.
├── .claude-plugin/
│   └── marketplace.json     # Claude Code marketplace (outlook + obsidian)
├── obsidian/                # Obsidian Claude Code plugin (CLI)
│   ├── .claude-plugin/plugin.json
│   └── .mcp.json            # MCP server registration (stdio)
├── obsidian-dxt/            # Obsidian Claude Desktop Extension (Cowork)
│   ├── manifest.json        # .mcpb manifest — edit paths + API key here
│   ├── server/index.js      # Node wrapper spawning mcp-server.exe
│   └── README.md            # Full setup guide
├── outlook/                 # the Outlook Claude Code plugin
│   ├── .claude-plugin/plugin.json
│   ├── .mcp.json            # MCP server registration using ${CLAUDE_PLUGIN_ROOT}
│   ├── commands/            # slash commands
│   ├── skills/outlook-assistant/SKILL.md
│   ├── src/outlook_mcp/     # Python MCP server
│   │   ├── outlook.py       # COM wrapper (pure, no MCP dep)
│   │   └── server.py        # FastMCP tool definitions
│   └── pyproject.toml
├── dxt/                     # Claude desktop extension staging
│   ├── manifest.json        # .mcpb manifest (Anthropic desktop extension spec)
│   └── server/              # populated by build-dxt.sh (gitignored)
├── scripts/
│   └── build-dxt.sh         # stages server code and packs outlook-<version>.mcpb
├── icon.jpg                 # source icon (auto-converted to PNG at build)
├── LICENSE
└── README.md
```

### Build the desktop extension

```bash
bash scripts/build-dxt.sh
# → outlook-0.1.0.mcpb in repo root
```

The script copies `outlook/src/outlook_mcp` and `outlook/pyproject.toml` into `dxt/server/`, converts `icon.jpg` → PNG via `uv`+Pillow, and packs the bundle with `@anthropic-ai/mcpb`. Distribute the `.mcpb` by attaching it to a GitHub Release.

The `outlook.py` COM wrapper has no MCP dependency and can be imported directly for scripting:

```python
from outlook_mcp import outlook as ol
ol.list_emails(folder="inbox", unread_only=True, limit=10)
```

## Tool reference

### Folder paths

Most tools accept a `folder` string:

| Input | Resolves to |
|---|---|
| omitted or `"inbox"` | Default Inbox |
| `"sent"`, `"drafts"`, `"deleted"`, `"junk"`, `"outbox"`, `"calendar"` | Default special folders |
| `"posteingang"`, `"entwürfe"`, `"kalender"`, … | German equivalents (see below) |
| `"inbox/Processed/Q1"` | Subfolder under default Inbox |
| `"account@example.com/Inbox/Processed"` | Walk from a specific store |

### Localized folder names

The following German folder aliases are supported alongside the English names:

| German | English | Folder |
|---|---|---|
| `posteingang` | `inbox` | Inbox |
| `gesendete elemente` / `gesendet` | `sent` | Sent Items |
| `entwürfe` | `drafts` | Drafts |
| `gelöschte elemente` / `papierkorb` | `deleted` / `trash` | Deleted Items |
| `junk-e-mail` | `junk` | Junk Email |
| `postausgang` | `outbox` | Outbox |
| `kalender` | `calendar` | Calendar |
| `kontakte` | `contacts` | Contacts |
| `aufgaben` | `tasks` | Tasks |

Subfolder resolution is case-insensitive, so paths like `"account@x.com/Inbox/Projects"` work even when Outlook displays the folder as "Posteingang" internally.

### Complete tool list

**Email:** `list_accounts`, `list_folders`, `list_emails`, `read_email`, `search_emails`, `move_email`, `delete_email`, `create_draft`, `update_draft`, `send_email`

**Calendar:** `list_calendar_events`, `create_calendar_event`, `update_calendar_event`, `delete_calendar_event`, `respond_to_invite`

**Categories:** `list_categories`, `set_email_categories`, `set_event_categories`

## Security & safety

- **No cloud auth.** The plugin talks to the Outlook COM object on your machine. If you can read the mailbox in Outlook, the plugin can too — nothing more, nothing less.
- **Programmatic-access prompt.** Older Outlook versions or some group-policy configurations display a warning when an external process sends mail. Modern Office 365 generally suppresses it for trusted executables. IT admins can configure via Trust Center → Programmatic Access.
- **Safe-by-default tool use.** The bundled `outlook-assistant` skill instructs Claude to draft before sending, soft-delete by default, and confirm batch operations.
- **Sending is real.** `send_email` sends. For reviewable workflows, prefer `create_draft` and let a human hit Send.
- **DASL filter injection prevention.** User input in search and filter parameters is escaped (single quotes, `%` and `_` LIKE wildcards) to prevent filter manipulation.
- **Sensitive file blocking.** Attachment paths are validated against a blocklist of sensitive patterns (`.ssh`, `.aws`, `.env`, `.pem`, `.key`, `credentials.json`, etc.) to prevent accidental credential exfiltration.
- **HTML sanitization.** Dangerous HTML tags (`<script>`, `<iframe>`, `<form>`, `<object>`, `<embed>`, `<applet>`, `<meta http-equiv>`) are stripped from `html_body` content before writing to drafts or sending, mitigating prompt injection attacks.

## Roadmap

- Publish to PyPI so `uvx outlook-mcp` works without cloning
- Contacts + tasks tools
- Attachment download (save to disk)
- C#/.NET port for single-exe distribution (easier IT whitelisting than a Python venv)

---

## Obsidian

Connects Claude to a local [Obsidian](https://obsidian.md) vault via the [MCP Tools plugin](https://github.com/jacksteamdev/obsidian-mcp-tools) (jacksteamdev). Supports reading, creating, editing and deleting notes, full-text search, and semantic search via Smart Connections.

### Prerequisites

- **Obsidian** with the **MCP Tools** plugin installed (click "Install Server" in plugin settings)
- **Local REST API** plugin (coddingtonbear) — must be active
- Optional: **Smart Connections** plugin for semantic search

### Install for Cowork (Claude Desktop Extension)

1. Copy `obsidian-dxt/` to `%APPDATA%\Claude\Claude Extensions\local.mcpb.<your-name>.obsidian\`
2. Edit `manifest.json` — set `MCP_SERVER_PATH` and `command` to your `mcp-server.exe` path, set `OBSIDIAN_API_KEY` to your key
3. Register in `%APPDATA%\Claude\extensions-installations.json`
4. Restart Claude Desktop, open a new Cowork conversation

See [`obsidian-dxt/README.md`](obsidian-dxt/README.md) for the full step-by-step guide.

### Install for Claude Code (CLI)

```powershell
claude mcp add obsidian \
  --env OBSIDIAN_API_KEY=<your-key> \
  -- "<vault-path>\.obsidian\plugins\mcp-tools\bin\mcp-server.exe"
```

Or install the plugin from this repo:

```powershell
claude plugin marketplace add andiba/outlook-extension-de
claude plugin install obsidian
```

### Obsidian tools

`search_vault`, `search_vault_simple`, `search_vault_smart`, `list_vault_files`, `get_vault_file`, `create_vault_file`, `append_to_vault_file`, `patch_vault_file`, `delete_vault_file`, `get_active_file`, `update_active_file`, `patch_active_file`, `append_to_active_file`, `delete_active_file`, `execute_template`, `show_file_in_obsidian`, `get_server_info`, `fetch`

---

## Author

**Gabriel Denny**
- Website: [www.gabrieldenny.com](https://www.gabrieldenny.com)
- LinkedIn: [gabrieljdenny](https://www.linkedin.com/in/gabrieljdenny)
- GitHub: [@gabrieldenny-del](https://github.com/gabrieldenny-del)

## License

MIT © 2026 Gabriel Denny — see [LICENSE](LICENSE).
