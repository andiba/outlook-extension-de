# Obsidian Vault - Claude Desktop Extension

Verbindet Claude Desktop (Cowork) mit deinem lokalen Obsidian-Vault via [MCP Tools Plugin](https://github.com/jacksteamdev/obsidian-mcp-tools).

## Voraussetzungen

- **Windows** (die Extension nutzt die native MCP-Server-Binary)
- **Obsidian** mit installiertem und aktiviertem **MCP Tools** Plugin (jacksteamdev)
- **Local REST API** Plugin (coddingtonbear) - muss laufen, damit der MCP-Server antworten kann
- Optional: **Smart Connections** Plugin fuer semantische Suche
- Optional: **[Git](https://github.com/Vinzent03/obsidian-git)** Plugin (Vinzent) fuer automatischen Vault-Sync via GitHub

## Setup

### 1. MCP Tools Plugin installieren

1. Obsidian oeffnen > Einstellungen > Community Plugins > "MCP Tools" suchen und installieren
2. Im Plugin auf **"Install Server"** klicken - das legt `mcp-server.exe` ab unter:
   ```
   <VAULT_PFAD>\.obsidian\plugins\mcp-tools\bin\mcp-server.exe
   ```
3. Den angezeigten **API Key** kopieren

### 2. manifest.json anpassen

In `manifest.json` die zwei Platzhalter ersetzen:

```json
"mcp_config": {
  "command": "<DEIN_VAULT_PFAD>\\.obsidian\\plugins\\mcp-tools\\bin\\mcp-server.exe",
  "env": {
    "MCP_SERVER_PATH": "<DEIN_VAULT_PFAD>\\.obsidian\\plugins\\mcp-tools\\bin\\mcp-server.exe",
    "OBSIDIAN_API_KEY": "<DEIN_API_KEY>"
  }
}
```

- `command` und `MCP_SERVER_PATH` - absoluter Pfad zur `mcp-server.exe` aus Schritt 1 (muss an beiden Stellen identisch sein)
- `OBSIDIAN_API_KEY` - API Key aus den MCP Tools Plugin-Einstellungen
- **Wichtig:** Das `command`-Feld darf nicht fehlen, sonst wird die Extension nicht geladen

### 3. Extension installieren

**Option A - Manuell kopieren (empfohlen):**

1. Den gesamten Ordner nach `%APPDATA%\Claude\Claude Extensions\` kopieren:
   ```
   %APPDATA%\Claude\Claude Extensions\local.mcpb.<dein-name>.obsidian\
   ├── manifest.json
   └── server\
       └── index.js
   ```
2. In `%APPDATA%\Claude\extensions-installations.json` einen Eintrag hinzufuegen (siehe bestehendes Format)
3. Claude Desktop komplett beenden und neu starten

**Option B - Als .mcpb packen:**

1. Die Dateien als ZIP packen (manifest.json im Root der ZIP)
2. ZIP-Endung zu `.mcpb` umbenennen
3. Doppelklick auf die `.mcpb`-Datei (Datei-Zuordnung funktioniert nicht immer - dann Option A nehmen)

### 4. Testen

1. Neue Konversation in Cowork oeffnen
2. Die MCP-Tools sollten als `mcp__Obsidian_Vault__*` sichtbar sein
3. Testen mit: "Liste alle Dateien in meinem Vault auf"

## Wichtig

- **Obsidian muss laufen** waehrend du Cowork benutzt
- **Local REST API** Plugin muss aktiv sein
- Nach Aenderungen an der Extension: Claude Desktop **komplett beenden** und neu starten
- Laufende Cowork-Sessions sehen neue Tools erst nach Neustart

## Verfuegbare Tools

| Tool | Beschreibung |
|------|-------------|
| `get_vault_file` | Datei aus dem Vault lesen |
| `list_vault_files` | Dateien im Vault auflisten |
| `search_vault` | Suche nach Dateiname oder Inhalt |
| `search_vault_simple` | Einfache Textsuche |
| `search_vault_smart` | Semantische Suche (braucht Smart Connections) |
| `create_vault_file` | Neue Datei erstellen |
| `patch_vault_file` | Datei aktualisieren |
| `append_to_vault_file` | An Datei anhaengen |
| `delete_vault_file` | Datei loeschen |
| `get_active_file` | Aktuell geoeffnete Datei lesen |
| `update_active_file` | Aktuell geoeffnete Datei ueberschreiben |
| `patch_active_file` | Aktuell geoeffnete Datei patchen |
| `append_to_active_file` | An aktuelle Datei anhaengen |
| `delete_active_file` | Aktuelle Datei loeschen |
| `show_file_in_obsidian` | Datei in Obsidian oeffnen |
| `execute_template` | Obsidian-Template ausfuehren |
| `get_server_info` | Server-Info abrufen |
| `fetch` | URL abrufen |

## Hintergrund

Warum eine Extension und nicht einfach ein Custom Connector in Cowork?

- Der MCP-Server ist **stdio-basiert** (keine HTTP-API)
- Cowork's Custom Connector lehnt **localhost-URLs** ab
- `mcp-proxy --tunnel` hat Routing-Probleme (404 auf `/sse`)
- Die Extension ist der einzige Weg, einen lokalen stdio-MCP-Server in Cowork einzubinden

## Vault-Sync (optional)

Fuer Sync auf mehreren Rechnern das **Git** Community Plugin (Vinzent) installieren:

- Auto-Commit alle 10 Minuten
- Auto-Push nach jedem Commit
- Auto-Pull beim Start

Vault als privates GitHub-Repo anlegen, `.gitignore` fuer `.obsidian/plugins/*/bin/`, `.smart-env/`, `workspace.json` und Plugin-`data.json` mit API-Keys.

## Links

- [MCP Tools Plugin](https://github.com/jacksteamdev/obsidian-mcp-tools)
- [Local REST API Plugin](https://github.com/coddingtonbear/obsidian-local-rest-api)
- [Smart Connections](https://github.com/brianpetro/obsidian-smart-connections)
- [Obsidian Git Plugin](https://github.com/Vinzent03/obsidian-git)
