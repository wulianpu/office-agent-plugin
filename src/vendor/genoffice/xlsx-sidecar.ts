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

interface PendingRequest {
  resolve: (v: unknown) => void;
  reject: (e: Error) => void;
  timer: NodeJS.Timeout;
}

const STDERR_RING_BYTES = 8 * 1024;
/** Default per-request deadline: a live-but-wedged sidecar must let the
 *  caller fall back to the JS renderer in seconds-to-tens-of-seconds,
 *  never block a preview indefinitely (round 9, P1-high). */
const DEFAULT_REQUEST_TIMEOUT_MS = 15_000;
/** Cooldown after a hard timeout: fail fast to the fallback path, then allow
 *  one fresh respawn attempt (bounded wedge cost, no permanent latch). */
const UNHEALTHY_COOLDOWN_MS = 30_000;

export class XlsxSidecarClient {
  private child?: ChildProcessWithoutNullStreams;
  private pending = new Map<string, PendingRequest>();
  private seq = 0;
  private startError?: Error;
  private unhealthyUntil = 0;
  private readonly stderrRing: string[] = [];

  constructor(
    readonly exePath: string = process.env.XLSX_SIDECAR_PATH ?? defaultSidecarPath(),
    private readonly options: { requestTimeoutMs?: number; unhealthyCooldownMs?: number } = {}
  ) {}

  get available(): boolean {
    return existsSync(this.exePath);
  }

  /** Bounded recent stderr for diagnostics — never an unbounded string. */
  recentStderr(): string {
    return this.stderrRing.join("");
  }

  /** In-flight request count (introspection for tests/telemetry). */
  pendingCount(): number {
    return this.pending.size;
  }

  private ensureStarted(): void {
    if (this.startError) throw this.startError;
    if (this.child) return;
    // Cooldown after a hard timeout: fail fast so previews use the JS
    // fallback immediately; a fresh spawn attempt is allowed afterwards.
    if (Date.now() < this.unhealthyUntil) {
      throw new Error(
        `xlsx sidecar cooling down after timeout (retry in ${Math.ceil((this.unhealthyUntil - Date.now()) / 1000)}s)`
      );
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
      this.child = this.startChild();
      // Drain stderr into a bounded ring: an unread pipe fills and the OS
      // backpressures the child into a live-but-wedged state. Never accrue
      // an unbounded string.
      this.child.stderr.on("data", (chunk: Buffer) => {
        this.stderrRing.push(chunk.toString("utf8"));
        let bytes = this.stderrRing.reduce((n, part) => n + part.length, 0);
        while (this.stderrRing.length > 1 && bytes > STDERR_RING_BYTES) {
          bytes -= this.stderrRing.shift()!.length;
        }
      });
      // A dead stdin (EPIPE after a crash) must settle pending requests —
      // child-level error/exit events alone are not guaranteed to fire first.
      this.child.stdin.on("error", (error) => {
        this.failAllPending(new Error(`sidecar stdin error: ${String(error)}`));
      });
      const rl = createInterface({ input: this.child.stdout! });
      rl.on("line", (line) => {
        if (!line.trim()) return;
        try {
          const response = JSON.parse(line) as { requestId: string; ok: boolean; result?: unknown; error?: { code: string; message: string } };
          const entry = this.pending.get(response.requestId);
          if (!entry) return;
          this.pending.delete(response.requestId);
          clearTimeout(entry.timer);
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
        this.failAllPending(this.startError);
        this.child = undefined;
      });
      this.child.on("exit", () => {
        this.failAllPending(new Error("sidecar exited"));
        this.child = undefined;
      });
    } catch (error) {
      this.startError = error instanceof Error ? error : new Error(String(error));
      throw this.startError;
    }
  }

  /** Spawn the native process — protected so deterministic tests substitute
   *  a fake child without a real executable (round 9). */
  protected startChild(): ChildProcessWithoutNullStreams {
    return spawn(this.exePath, [], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true }) as ChildProcessWithoutNullStreams;
  }

  private async request<T>(command: string, payload: Record<string, unknown>): Promise<T> {
    this.ensureStarted();
    const requestId = `req-${++this.seq}-${randomUUID().slice(0, 8)}`;
    const envelope = { version: PROTOCOL_VERSION, requestId, command, ...payload };
    const timeoutMs = this.options.requestTimeoutMs ?? DEFAULT_REQUEST_TIMEOUT_MS;
    return new Promise<T>((resolvePromise, rejectPromise) => {
      const entry = {
        resolve: (value: unknown) => {
          clearTimeout(entry.timer);
          resolvePromise(value as T);
        },
        reject: (error: Error) => {
          clearTimeout(entry.timer);
          rejectPromise(error);
        },
        timer: setTimeout(() => {
          this.pending.delete(requestId);
          // Live-but-wedged: kill the process ('exit' settles everything
          // still pending) and latch a cooldown so subsequent previews fail
          // fast into the JS renderer instead of paying another timeout.
          this.unhealthyUntil =
            Date.now() + (this.options.unhealthyCooldownMs ?? UNHEALTHY_COOLDOWN_MS);
          entry.reject(new Error(`sidecar request timeout after ${timeoutMs}ms: ${command}`));
          this.child?.kill();
        }, timeoutMs) as NodeJS.Timeout
      };
      this.pending.set(requestId, entry);
      try {
        this.child!.stdin.write(`${JSON.stringify(envelope)}\n`);
      } catch (error) {
        this.pending.delete(requestId);
        clearTimeout(entry.timer);
        rejectPromise(new Error(`sidecar stdin write failed: ${String(error)}`));
      }
    });
  }

  private failAllPending(error: Error): void {
    for (const [, entry] of this.pending) {
      clearTimeout(entry.timer);
      entry.reject(error);
    }
    this.pending.clear();
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
    this.failAllPending(new Error("sidecar disposed"));
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
 * open → sheet viewport(s) → close. Rejects when the sidecar is unavailable —
 * callers fall back to the in-process renderer. Round 10: `sheet` selects a
 * named sheet; `range` (1-based, inclusive) also carries the window ORIGIN —
 * `Data!D10:F30` must read from D10, not merely resize the top-left window.
 */
export async function sidecarPreviewWindow(
  client: XlsxSidecarClient,
  path: string,
  options: {
    sheet?: string;
    /** 1-based inclusive window origin+extent; converted to the native
     *  0-based protocol here — the only place the bases meet. */
    range?: { fromRow: number; toRow: number; fromCol: number; toCol: number };
    maxRows?: number;
    maxCols?: number;
    maxSheets?: number;
  } = {}
): Promise<Array<{ name: string; window: string[][]; rowCount?: number }>> {
  const maxRows = options.maxRows ?? 40;
  const maxCols = options.maxCols ?? 16;
  const startRow0 = options.range ? Math.max(0, Math.floor(options.range.fromRow) - 1) : 0;
  const startCol0 = options.range ? Math.max(0, Math.floor(options.range.fromCol) - 1) : 0;
  const opened = await client.open(path);
  const out: Array<{ name: string; window: string[][]; rowCount?: number }> = [];
  const sheets = options.sheet
    ? [
        ...(opened.sheets.find((s) => s.name === options.sheet)
          ? [opened.sheets.find((s) => s.name === options.sheet)!]
          : []),
        ...opened.sheets.filter((s) => s.name !== options.sheet)
      ]
    : opened.sheets;
  try {
    for (const sheet of sheets.slice(0, options.maxSheets ?? 4)) {
      // Clamp the WINDOW SIZE to the sheet's declared extent (fall back to a
      // probe window when metadata omits dimensions — read_range rejects
      // out-of-sheet reads); the origin offsets stay as requested.
      const rows = Math.max(1, Math.min(maxRows, sheet.rowCount ?? maxRows));
      const cols = Math.max(1, Math.min(maxCols, sheet.columnCount ?? maxCols));
      let result = await client
        .readRange(opened.sessionId, sheet.id, {
          startRow: startRow0,
          endRow: startRow0 + rows - 1,
          startColumn: startCol0,
          endColumn: startCol0 + cols - 1
        })
        .catch(async () =>
          client.readRange(opened.sessionId, sheet.id, {
            startRow: startRow0,
            endRow: startRow0 + 4,
            startColumn: startCol0,
            endColumn: startCol0 + 4
          })
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
