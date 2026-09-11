/**
 * Optional Electron wrapper (§7 process model): launches the Office Plugin
 * runtime and points a BrowserWindow at the localhost host URL. The web host
 * (src/host/preview-server.ts) is the actual renderer surface — this file is
 * the ~60-line desktop shell around it, matching the design's "Harness Main
 * manages Window/Process; Renderer hosts the editor" split.
 *
 * Run:  npm i -D electron && npx electron tools/host/electron.mjs <files...>
 */

import { spawn } from "node:child_process";
import { app, BrowserWindow } from "electron";

const files = process.argv.slice(2);

async function main() {
  // The runtime+host run as a child Node process (mirrors §7: main, renderer,
  // runtime service as separate processes).
  const host = spawn(process.execPath, ["dist/host/main.js", ...files], {
    stdio: ["ignore", "pipe", "inherit"],
    windowsHide: true
  });
  let url = "";
  host.stdout.on("data", (chunk: Buffer) => {
    const text = chunk.toString("utf8");
    process.stdout.write(text);
    const match = text.match(/Office Plugin host: (http:\/\/\S+)/);
    if (match) {
      url = match[1]!;
      void openWindow();
    }
  });

  async function openWindow() {
    const win = new BrowserWindow({
      width: 1280,
      height: 800,
      title: "Office Plugin",
      autoHideMenuBar: true
    });
    await win.loadURL(url);
  }

  app.on("window-all-closed", () => {
    host.kill();
    app.quit();
  });
  app.on("before-quit", () => host.kill());
}

app.whenReady().then(() => void main().catch((error) => {
  console.error(error);
  app.quit();
}));
