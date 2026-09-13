/**
 * XLSX Rust sidecar client (§30, §146): spawns the vendored GenOffice
 * xlsx-sidecar binary (built via cargo) and speaks its line-delimited JSON
 * protocol. The sidecar keeps workbook memory out of this process — the
 * renderer working set stays bounded by the viewport, not the file (§31).
 *
 * Protocol (from apps/sheets/native/xlsx-engine/src/main.rs, v1):
 *   request  {version, requestId, command: "open"|"read_range"|"close"|...}
 *   response {version, requestId, ok, result|error}  (camelCase fields)
 */

import { spawn, type ChildProcessWithoutNullStreams } from "node:child_process";
import { createInterface } from "node:readline";
import { existsSync } from "node:fs";
import { join, resolve } from "node:path";
import { randomUUID } from "node:crypto";

const PROTOCOL_VERSION = 1;

export interface SidecarSheetInfo {
  id: string;
  name: string;
  rowCount?: number;
  columnCount?: number;
  [key: string]: unknown;
}

export interface SidecarOpenResult {
  sessionId: string;
  sheets: SidecarSheetInfo[];
  [key: string]: unknown;
}

export interface SidecarRangeResult {
  /** Loosely typed: the exact cell payload shape is engine-owned. */
  cells?: unknown;
  rows?: unknown;
  [key: string]: unknown;
}

export function defaultSidecarPath(): string {
  return resolve(
    join(
      "vendor",
      "genoffice",
      "apps",
      "sheets",
      "native",
      "xlsx-engine",
      "target",
      "release",
      process.platform === "win32" ? "xlsx-sidecar.exe" : "xlsx-sidecar"
    )
  );
}

export class XlsxSidecarClient {
  private child?: ChildProcessWithoutNullStreams;
  private pending = new Map<string, { resolve: (v: unknown) => void; reject: (e: Error) => void }>();
  private seq = 0;
  private startError?: Error;

  constructor(readonly exePath: string = process.env.XLSX_SIDECAR_PATH ?? defaultSidecarPath()) {}

  get available(): boolean {
    return existsSync(this.exePath);
  }

  private ensureStarted(): void {
    if (this.child || this.startError) {
      if (this.startError) throw this.startError;
      return;
    }
    // Fast-fail before spawning: a missing binary must reject the FIRST
    // request immediately — spawn() reports ENOENT asynchronously via the
    // 'error' event, and an unobserved one leaves every pending request
    // unsettled forever (the CI integration red build hung exactly there).
    if (!existsSync(this.exePath)) {
      this.startError = new Error(`xlsx sidecar binary not found: ${this.exePath}`);
      throw this.startError;
    }
    try {
      this.child = spawn(this.exePath, [], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true }) as ChildProcessWithoutNullStreams;
      const rl = createInterface({ input: this.child.stdout! });
      rl.on("line", (line) => {
        if (!line.trim()) return;
        try {
          const response = JSON.parse(line) as { requestId: string; ok: boolean; result?: unknown; error?: { code: string; message: string } };
          const entry = this.pending.get(response.requestId);
          if (!entry) return;
          this.pending.delete(response.requestId);
          if (response.ok) entry.resolve(response.result);
          else entry.reject(new Error(`sidecar ${response.error?.code}: ${response.error?.message}`));
        } catch {
          // Ignore malformed lines (engine stderr diagnostics stay on stderr).
        }
      });
      this.child.on("error", (error) => {
        // Spawn/early-lifecycle failure: settle every pending request NOW and
        // latch the error so later calls fail fast instead of hanging.
        this.startError = error instanceof Error ? error : new Error(String(error));
        for (const [, entry] of this.pending) entry.reject(this.startError);
        this.pending.clear();
        this.child = undefined;
      });
      this.child.on("exit", () => {
        for (const [, entry] of this.pending) entry.reject(new Error("sidecar exited"));
        this.pending.clear();
        this.child = undefined;
      });
    } catch (error) {
      this.startError = error instanceof Error ? error : new Error(String(error));
      throw this.startError;
    }
  }

  private request<T>(command: string, payload: Record<string, unknown>): Promise<T> {
    this.ensureStarted();
    const requestId = `req-${++this.seq}-${randomUUID().slice(0, 8)}`;
    const envelope = { version: PROTOCOL_VERSION, requestId, command, ...payload };
    return new Promise<T>((resolvePromise, rejectPromise) => {
      this.pending.set(requestId, { resolve: resolvePromise as (v: unknown) => void, reject: rejectPromise });
      this.child!.stdin.write(`${JSON.stringify(envelope)}\n`);
    });
  }

  open(path: string, locale = "en_US"): Promise<SidecarOpenResult> {
    return this.request("open", { path: resolve(path), locale });
  }

  readRange(sessionId: string, sheetId: string, range: { startRow: number; endRow: number; startColumn: number; endColumn: number }): Promise<SidecarRangeResult> {
    return this.request("read_range", { sessionId, sheetId, range });
  }

  close(sessionId: string): Promise<void> {
    return this.request("close", { sessionId });
  }

  async dispose(): Promise<void> {
    const child = this.child;
    this.child = undefined;
    if (!child) return;
    child.stdin.end();
    child.kill();
    await new Promise((resolveExit) => {
      child.once("exit", resolveExit);
      setTimeout(resolveExit, 2000);
    });
  }
}

/**
 * Bounded preview window through the sidecar (§30 read path letter):
 * open → first sheet viewport → close. Rejects when the sidecar is
 * unavailable — callers fall back to the in-process renderer.
 */
export async function sidecarPreviewWindow(
  client: XlsxSidecarClient,
  path: string,
  options: { maxRows?: number; maxCols?: number; maxSheets?: number } = {}
): Promise<Array<{ name: string; window: string[][]; rowCount?: number }>> {
  const maxRows = options.maxRows ?? 40;
  const maxCols = options.maxCols ?? 16;
  const opened = await client.open(path);
  const out: Array<{ name: string; window: string[][]; rowCount?: number }> = [];
  try {
    for (const sheet of opened.sheets.slice(0, options.maxSheets ?? 4)) {
      // Clamp to the sheet's declared extent; fall back to a probe window
      // when metadata omits dimensions (read_range rejects out-of-sheet).
      const rows = Math.max(1, Math.min(maxRows, sheet.rowCount ?? maxRows));
      const cols = Math.max(1, Math.min(maxCols, sheet.columnCount ?? maxCols));
      let result = await client
        .readRange(opened.sessionId, sheet.id, { startRow: 0, endRow: rows - 1, startColumn: 0, endColumn: cols - 1 })
        .catch(async () =>
          client.readRange(opened.sessionId, sheet.id, { startRow: 0, endRow: 4, startColumn: 0, endColumn: 4 })
        );
      result = await result;
      out.push({ name: sheet.name, window: normalizeWindow(result, maxRows, maxCols), rowCount: sheet.rowCount });
    }
  } finally {
    await client.close(opened.sessionId).catch(() => undefined);
  }
  return out;
}

/** Scatter the engine's sparse cell records into a dense 2D string window. */
function normalizeWindow(result: SidecarRangeResult, maxRows: number, maxCols: number): string[][] {
  const grid: string[][] = [];
  const ensure = (row: number) => {
    while (grid.length <= row) grid.push([]);
    return grid[row]!;
  };
  if (Array.isArray(result.cells)) {
    for (const record of result.cells as Array<{ row?: number; column?: number; value?: unknown }>) {
      const row = record.row ?? 0;
      const col = record.column ?? 0;
      if (row >= maxRows || col >= maxCols) continue;
      const line = ensure(row);
      while (line.length <= col) line.push("");
      line[col] = record.value === null || record.value === undefined ? "" : String(record.value);
    }
  }
  return grid;
}
