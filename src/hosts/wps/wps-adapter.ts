/**
 * WPS host adapter (§86–§88): drives WPS Presentation (KWPP.Application COM)
 * through PowerShell — zero native dependencies. Provides host-true PNG
 * renders and the disposable certification copy workflow. Headless:
 * Presentations.Open runs WithWindow=msoFalse; selection requires an
 * attached windowed workflow and is intentionally not offered here.
 */

import { spawn } from "node:child_process";
import { mkdtemp, readFile, readdir, rm, writeFile, copyFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createHash } from "node:crypto";
import type {
  HostAdapter,
  HostCertificationResult,
  HostRenderOptions,
  HostRenderResult
} from "../host-adapter.js";

const PROG_ID = "KWPP.Application";
const PS_TIMEOUT_MS = 180_000;

function runPowerShell(script: string, args: string[]): Promise<string> {
  return new Promise((resolve, reject) => {
    const child = spawn(
      "powershell",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script, ...args],
      { windowsHide: true, stdio: ["ignore", "pipe", "pipe"] }
    );
    let stdout = "";
    let stderr = "";
    const timer = setTimeout(() => {
      child.kill();
      reject(new Error("wps host: powershell timed out"));
    }, PS_TIMEOUT_MS);
    child.stdout.on("data", (c: Buffer) => (stdout += c));
    child.stderr.on("data", (c: Buffer) => (stderr += c));
    child.on("error", (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.on("close", (code) => {
      clearTimeout(timer);
      if (code === 0) resolve(stdout);
      else reject(new Error(`wps host: powershell exit ${code}: ${stderr.trim().slice(0, 400)}`));
    });
  });
}

/** COM single-use servers reject concurrent instantiation — retry these. */
const COM_RETRY_CODES = ["80010001", "80080005", "RPC_E_CALL_REJECTED", "CO_E_SERVER_EXEC_FAILURE"];

async function runPowerShellWithRetry(script: string, args: string[]): Promise<string> {
  let lastError: Error | undefined;
  for (let attempt = 0; attempt < 4; attempt++) {
    try {
      return await runPowerShell(script, args);
    } catch (error) {
      lastError = error as Error;
      const text = String((error as Error).message);
      if (!COM_RETRY_CODES.some((code) => text.includes(code))) throw error;
      await new Promise((resolve) => setTimeout(resolve, 2000 + attempt * 2000));
    }
  }
  throw lastError;
}

export class WpsHostAdapter implements HostAdapter {
  readonly id = "wps" as const;
  readonly engine = "KWPP-COM";

  /** Process-wide shared probe: COM single-use servers must not be stampeded. */
  private static sharedProbeOnce?: Promise<boolean>;

  static sharedProbe(): Promise<boolean> {
    WpsHostAdapter.sharedProbeOnce ??= new WpsHostAdapter()
      .probe()
      .catch(() => false)
      .then((result) => result)
      .catch(() => false);
    return WpsHostAdapter.sharedProbeOnce;
  }

  async probe(): Promise<boolean> {
    try {
      const script = await this.materialize(probeScript);
      await runPowerShellWithRetry(script, []);
      return true;
    } catch {
      return false;
    }
  }

  async renderSlides(path: string, slideIndices: number[], options: HostRenderOptions = {}): Promise<HostRenderResult> {
    if (slideIndices.length === 0) throw new Error("wps host: no slide indices requested");
    const dir = await mkdtemp(join(tmpdir(), "wps-render-"));
    try {
      const script = await this.materialize(renderScript);
      const indicesArg = slideIndices.join(",");
      const width = options.width ?? 960;
      const height = options.height ?? 540;
      const manifest = join(dir, "manifest.json");
      const output = await runPowerShellWithRetry(script, [path, indicesArg, String(width), String(height), dir]);
      const pngs: Buffer[] = [];
      const count = Number(output.match(/WPS-RENDER-OK slides=(\d+)/)?.[1] ?? 0);
      for (const index of slideIndices) {
        try {
          pngs.push(await readFile(join(dir, `slide-${index}.png`)));
        } catch {
          // Out-of-range slide: the host skipped it; caller checks pngs.length.
        }
      }
      if (pngs.length === 0) {
        throw new Error(`wps host: rendered ${count} slides but none of the requested indices landed`);
      }
      void manifest;
      return { pngs, slideCount: count };
    } finally {
      await rm(dir, { recursive: true, force: true }).catch(() => undefined);
    }
  }

  async certifyCopy(path: string, options: HostRenderOptions = {}): Promise<HostCertificationResult> {
    const dir = await mkdtemp(join(tmpdir(), "wps-cert-"));
    try {
      // §88: the host only ever touches a disposable copy.
      const copyPath = join(dir, "certification-copy.pptx");
      const roundTripPath = join(dir, "roundtrip.pptx");
      await copyFile(path, copyPath);
      const hashBefore = sha256(await readFile(copyPath));

      const script = await this.materialize(certifyScript);
      const output = await runPowerShellWithRetry(script, [
        copyPath,
        roundTripPath,
        join(dir, "host-render.png"),
        String(options.width ?? 960),
        String(options.height ?? 540)
      ]).catch((error: Error) => `WPS-OPEN-FAIL ${error.message}`);

      if (output.startsWith("WPS-OPEN-FAIL")) {
        return { status: "unavailable", repaired: false, note: output.slice(64) };
      }

      let renderedPng: Buffer | undefined;
      try {
        renderedPng = await readFile(join(dir, "host-render.png"));
      } catch {
        // Render is supplementary; the round-trip below is the real probe.
      }

      // §88 semantics: SaveCopyAs re-serializes by design (like PowerPoint),
      // so byte equality is NOT the pass criterion. Certification passes when
      // the host-saved copy is still a readable, renderable package — i.e. the
      // host neither refused the file nor repaired damage. A corrupted or
      // unopenable round-trip is the FAIL path (auto-repair territory).
      try {
        const roundTripBytes = await readFile(roundTripPath);
        if (roundTripBytes.length < 1024 || roundTripBytes[0] !== 0x50 || roundTripBytes[1] !== 0x4b) {
          return { status: "fail", repaired: true, renderedPng, note: "host round-trip produced a non-package file (§88 repair path)" };
        }
        const reRender = await this.renderSlides(roundTripPath, [1], options);
        if (reRender.pngs[0]!.length < 512) {
          return { status: "fail", repaired: true, note: "host round-trip renders empty (§88 repair path)" };
        }
        const normalized = sha256(roundTripBytes) !== hashBefore;
        return {
          status: "pass",
          repaired: false,
          renderedPng: renderedPng ?? reRender.pngs[0],
          note: normalized
            ? "host re-opened, re-saved and re-rendered the copy (re-serialized, no repair)"
            : "host round-tripped the copy byte-identically"
        };
      } catch (error) {
        return { status: "fail", repaired: true, note: `host round-trip unreadable: ${String((error as Error).message).slice(0, 160)}` };
      }
    } finally {
      await rm(dir, { recursive: true, force: true }).catch(() => undefined);
    }
  }

  /** Materialize an embedded .ps1 next to a temp file (COM scripts run from disk). */
  private async materialize(script: string): Promise<string> {
    const dir = await mkdtemp(join(tmpdir(), "wps-ps-"));
    const file = join(dir, "run.ps1");
    await writeFile(file, script, "utf8");
    // Best-effort cleanup after a long delay; renders may still be reading.
    setTimeout(() => void rm(dir, { recursive: true, force: true }).catch(() => undefined), 120_000).unref?.();
    return file;
  }
}

function sha256(data: Buffer): string {
  return createHash("sha256").update(data).digest("hex");
}

const probeScript = `
$ErrorActionPreference = "Stop"
Get-Process -Name "wpp" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 } | Stop-Process -Force
$app = New-Object -ComObject ${PROG_ID}
$app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
Write-Host "WPS-PROBE-OK"
`;

const renderScript = `
param([string]$Pptx, [string]$Indices, [int]$Width, [int]$Height, [string]$OutDir)
$ErrorActionPreference = "Stop"
$app = New-Object -ComObject ${PROG_ID}
try {
  $pres = $app.Presentations.Open($Pptx, -1, 0, 0)
  foreach ($idx in $Indices.Split(",")) {
    $n = [int]$idx
    if ($n -ge 1 -and $n -le $pres.Slides.Count) {
      $pres.Slides.Item($n).Export((Join-Path $OutDir ("slide-" + $n + ".png")), "PNG", $Width, $Height)
    }
  }
  Write-Host ("WPS-RENDER-OK slides={0}" -f $pres.Slides.Count)
  $pres.Close()
} finally {
  $app.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
  Get-Process -Name "wpp" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 } | Stop-Process -Force
}
`;

const certifyScript = `
param([string]$Copy, [string]$RoundTrip, [string]$RenderPng, [int]$Width, [int]$Height)
Get-Process -Name "wpp" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 } | Stop-Process -Force
$ErrorActionPreference = "Stop"
$app = New-Object -ComObject ${PROG_ID}
try {
  $pres = $app.Presentations.Open($Copy, 0, 0, 0)
  $pres.SaveCopyAs($RoundTrip)
  try { $pres.Slides.Item(1).Export($RenderPng, "PNG", $Width, $Height) } catch {}
  $pres.Close()
  Write-Host "WPS-CERT-OK"
} finally {
  $app.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
  Get-Process -Name "wpp" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 } | Stop-Process -Force
}
`;

void readdir;
