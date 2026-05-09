#!/usr/bin/env bash
# Build the Claude Desktop Extension (.mcpb) bundle for Outlook.
#
# Prerequisites:
#   - bash (Git Bash / WSL on Windows; native on macOS / Linux)
#   - python3 on PATH (used to read the manifest version reliably)
#   - Node.js + npx (for @anthropic-ai/mcpb)
#   - uv (for on-demand Pillow to convert the icon)
#
# Output: outlook-<version>.mcpb in the repo root.

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$REPO_ROOT"

# Use python's json parser instead of grep+sed: macOS BSD sed does not
# support \s, and a regex over the manifest is also fragile because the
# substring "version" appears inside the multi-line "long_description"
# field. python3 ships with macOS and is available everywhere bash runs.
VERSION=$(python3 -c 'import json; print(json.load(open("dxt/manifest.json"))["version"])')
OUTPUT="outlook-${VERSION}.mcpb"

echo "==> Cleaning bundle"
rm -rf dxt/server dxt/icon.*
mkdir -p dxt/server/src

echo "==> Staging server code into dxt/server/"
cp outlook/pyproject.toml dxt/server/pyproject.toml
cp -r outlook/src/outlook_mcp dxt/server/src/outlook_mcp
# Remove __pycache__ copied from live dev tree
find dxt/server -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null || true

echo "==> Preparing icon (must be PNG for .mcpb)"
if [[ -f icon.png ]]; then
  cp icon.png dxt/icon.png
elif [[ -f icon.jpg ]]; then
  echo "    converting icon.jpg -> dxt/icon.png via uv + pillow"
  uv run --with pillow python -c "from PIL import Image; Image.open('icon.jpg').convert('RGBA').save('dxt/icon.png')"
else
  echo "    no icon found; bundle will have no icon"
fi

echo "==> Packing bundle as ${OUTPUT}"
npx --yes @anthropic-ai/mcpb pack dxt "${OUTPUT}"

echo ""
echo "Built: ${OUTPUT}"
ls -la "${OUTPUT}"
