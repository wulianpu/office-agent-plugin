/**
 * MCP stdio launcher: `office-mcp` binary entry. Hosts spawn this process to
 * expose the office.* tool surface over newline-delimited JSON-RPC (§P8).
 */

import { OfficePlugin } from "../plugin/office-plugin.js";
import { McpStdioServer } from "./server.js";

export async function launchMcpServer(options: { workspaceRoot?: string } = {}): Promise<void> {
  const plugin = await OfficePlugin.create({ workspaceRoot: options.workspaceRoot });
  const server = new McpStdioServer(plugin.mcpTools);
  server.start(process.stdin, process.stdout);

  const shutdown = async () => {
    await plugin.dispose().catch(() => undefined);
    process.exit(0);
  };
  process.stdin.on("close", () => void shutdown());
  process.on("SIGINT", () => void shutdown());
  process.on("SIGTERM", () => void shutdown());
}

if (process.argv[1] && process.argv[1].endsWith("launch.js")) {
  launchMcpServer().catch((error) => {
    console.error(error);
    process.exit(1);
  });
}
