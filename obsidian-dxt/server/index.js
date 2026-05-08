// Wrapper: spawns the Obsidian MCP Tools server exe and pipes stdio
const { spawn } = require("child_process");

const exe = process.env.MCP_SERVER_PATH;
if (!exe) {
  console.error(
    "ERROR: MCP_SERVER_PATH not set.\n" +
    "Set it in manifest.json → server.mcp_config.env.MCP_SERVER_PATH\n" +
    "pointing to your mcp-server.exe from the MCP Tools Obsidian plugin."
  );
  process.exit(1);
}

const child = spawn(exe, [], {
  stdio: "inherit",
  env: { ...process.env },
});

child.on("error", (err) => {
  console.error(`Failed to start MCP server: ${err.message}`);
  console.error(`Path: ${exe}`);
  process.exit(1);
});

child.on("exit", (code) => process.exit(code ?? 1));
