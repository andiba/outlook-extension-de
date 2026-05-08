# CLAUDE.md

## Was ist dieses Repo?

Lokale Claude-Erweiterungen (Extensions + Plugins) fuer Windows, die Claude mit Desktop-Anwendungen verbinden — ohne Cloud-APIs oder OAuth.

**Repo:** `andiba/outlook-extension-de` auf GitHub

## Enthaltene Erweiterungen

### Outlook (`outlook/` + `dxt/`)
- **Claude Code Plugin** (`outlook/`): MCP-Server (Python/COM), Skill `outlook-assistant`, Slash-Commands (`/triage-inbox`, `/daily-digest`, `/draft-reply`, `/clean-inbox`)
- **Cowork Extension** (`dxt/`): `.mcpb`-Bundle, Build via `scripts/build-dxt.sh`
- Voraussetzungen: Windows, Outlook Desktop, `uv`

### Obsidian (`obsidian/` + `obsidian-dxt/`)
- **Claude Code Plugin** (`obsidian/`): stdio-MCP via `claude mcp add`
- **Cowork Extension** (`obsidian-dxt/`): Manuell nach `%APPDATA%\Claude\Claude Extensions\` kopieren, Pfade + API-Key in `manifest.json` anpassen
- Voraussetzungen: Windows, Obsidian mit MCP Tools Plugin (jacksteamdev), Local REST API Plugin
- Setup-Anleitung: `obsidian-dxt/README.md`

## Wichtige Pfade

- **Installierte Extensions:** `%APPDATA%\Claude\Claude Extensions\`
- **Extension-Registry:** `%APPDATA%\Claude\extensions-installations.json`
- **Obsidian Vault:** `C:\Entwicklung\obsidian\zweitesGehirn\zweitesGehirn\`
- **MCP-Server-Binary:** `<vault>\.obsidian\plugins\mcp-tools\bin\mcp-server.exe`

## Obsidian Vault Repo

- **GitHub:** [andiba/obsidian](https://github.com/andiba/obsidian.git) (privat)
- **Lokal:** `C:\Entwicklung\obsidian\zweitesGehirn\zweitesGehirn\`
- **Sync:** Obsidian Git Plugin (Vinzent) — auto-commit + push
- **Session-Routine:**
  1. Am Anfang jeder Session: `git pull` im Vault-Repo ausfuehren, damit Aenderungen von anderen Rechnern da sind
  2. Nach Aenderungen am Vault (Notizen erstellt/bearbeitet): User fragen ob committed und gepusht werden soll

## Hinweise

- Claude Code und Cowork haben **getrennte Plugin-Systeme** — Extensions muessen separat installiert werden
- Cowork laedt lokale MCP-Server nur ueber das Extension-System (`.mcpb`), nicht ueber `claude_desktop_config.json`
- Cowork Custom Connector akzeptiert keine localhost-URLs
- Nach Extension-Aenderungen: Claude Desktop komplett beenden + neu starten
- **Vault-Repo nicht vergessen:** Bei Vault-Aenderungen immer nach Commit+Push fragen
