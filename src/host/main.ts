/**
 * Host entry: `npm run host -- <files...>` — registers artifacts, serves the
 * preview UI on localhost, stays alive until Ctrl+C. This IS the renderer
 * container for headless environments; point an Electron BrowserWindow (or
 * any browser) at the printed URL for the full interactive surface.
 */

import { OfficePlugin } from "../plugin/office-plugin.js";
import { PreviewHost } from "./preview-server.js";

async function main(): Promise<void> {
  const files = process.argv.slice(2);
  if (files.length === 0) {
    console.error("usage: office-host <file.docx|xlsx|pptx> [more files...]");
    process.exit(2);
  }
  const plugin = await OfficePlugin.create({ workspaceRoot: ".office-runtime" });
  const host = new PreviewHost(plugin);
  const port = await host.start();
  for (const file of files) {
    try {
      const ref = await plugin.registerArtifact(file);
      console.log(`registered ${file} → ${ref}`);
    } catch (error) {
      console.error(`skip ${file}: ${(error as Error).message}`);
    }
  }
  console.log(`\nOffice Plugin host: ${host.url()}\n`);
  const shutdown = async () => {
    await host.stop();
    await plugin.dispose();
    process.exit(0);
  };
  process.on("SIGINT", () => void shutdown());
  process.on("SIGTERM", () => void shutdown());
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
