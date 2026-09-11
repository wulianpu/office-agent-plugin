/**
 * OfficeCliAdapter (§56, §63–§64): the only bridge to the OfficeCLI engine.
 *
 * Execution modes follow the design table (§63):
 *  - standalone: one batch = one open/execute/save cycle (atomic rollback)
 *  - resident: open → (batch in memory)* → save (read visibility barrier) →
 *    close (writer handoff barrier)
 *
 * On Windows the npm .cmd shim is resolved to its JS entry and executed with
 * the current process's Node binary — no shell quoting, no .cmd spawn
 * restrictions. OFFICECLI_BIN / OFFICECLI_ENTRY env vars override discovery.
 */

import { spawn } from "node:child_process";
import { access, constants, readFile } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";

export interface OfficeCliJson {
  success: boolean;
  data?: unknown;
  message?: string;
  error?: { error: string; code?: string; suggestion?: string };
}

export class OfficeCliError extends Error {
  constructor(
    readonly code: string,
    message: string,
    readonly suggestion?: string,
    readonly stdout?: string
  ) {
    super(message);
    this.name = "OfficeCliError";
  }
}

interface ResolvedCommand {
  command: string;
  baseArgs: string[];
}

export class OfficeCliAdapter {
  private resolved?: ResolvedCommand;
  private resolutionPromise?: Promise<ResolvedCommand>;

  constructor(readonly options: { timeoutMs?: number } = {}) {}

  /**
   * Standalone batch: ONE open/execute/save cycle with atomic rollback —
   * for one-shot tasks (§63). No resident involved.
   */
  async runBatchStandalone(file: string, items: unknown[]): Promise<{ results: Array<{ index: number; success: boolean; output: string }>; summary: Record<string, number> }> {
    const json = await this.exec(["batch", file, "--json"], {
      stdin: JSON.stringify(items)
    });
    const data = json.data as { results?: unknown; summary?: Record<string, number> } | undefined;
    return {
      results: (Array.isArray(data?.results) ? data!.results : []) as Array<{ index: number; success: boolean; output: string }>,
      summary: data?.summary ?? {}
    };
  }

  /**
   * Resident batch: items apply IN MEMORY through the live resident opened
   * by `open()` — disk visibility deferred to save/close (§64). Explicitly
   * named so callers never rely on implicit engine-side routing.
   */
  async runBatchResident(
    file: string,
    items: unknown[]
  ): Promise<{ results: Array<{ index: number; success: boolean; output: string }>; summary: Record<string, number> }> {
    return this.runBatchStandalone(file, items);
  }

  async open(file: string): Promise<void> {
    await this.exec(["open", file, "--json"]);
  }

  /** Read visibility barrier (§64, INV-06): flush in-memory mutation to disk, keep resident. */
  async save(file: string): Promise<void> {
    await this.exec(["save", file, "--json"]);
  }

  /** Writer handoff barrier (§64, INV-07): flush + release the resident. */
  async close(file: string): Promise<void> {
    await this.exec(["close", file, "--json"]);
  }

  async get(file: string, path: string): Promise<OfficeCliJson["data"]> {
    const json = await this.exec(["get", file, path, "--json"]);
    await this.releaseAutoResident(file);
    return json.data;
  }

  async query(file: string, selector: string): Promise<OfficeCliJson["data"]> {
    const json = await this.exec(["query", file, selector, "--json"]);
    await this.releaseAutoResident(file);
    return json.data;
  }

  async validate(file: string): Promise<{ passed: boolean; message: string }> {
    const json = await this.exec(["validate", file, "--json"]);
    await this.releaseAutoResident(file);
    const message = typeof json.data === "string" ? json.data : (json.message ?? "");
    // officecli signals validation failure with success:false + nonzero exit;
    // the message text itself may contain the word "errors" on the PASS path.
    return { passed: json.success, message };
  }

  /**
   * officecli read commands may asynchronously start a resident daemon that
   * holds a Windows file lock. A no-op close releases it; pool-owned
   * residents simply get reopened by the next mutation round.
   */
  private async releaseAutoResident(file: string): Promise<void> {
    await this.exec(["close", file, "--json"]).catch(() => undefined);
  }

  async version_(): Promise<string> {
    const json = await this.exec(["--version"]);
    return String(json.data ?? json.message ?? "").trim();
  }

  /** Generic command execution (fixture creation, host-specific probes). */
  run(args: string[], options: { stdin?: string } = {}): Promise<OfficeCliJson> {
    return this.exec(args, options);
  }

  private async exec(args: string[], options: { stdin?: string } = {}): Promise<OfficeCliJson> {
    const resolvedCommand = await this.resolve();
    const fullArgs = [...resolvedCommand.baseArgs, ...args];
    return new Promise<OfficeCliJson>((resolvePromise, rejectPromise) => {
      const child = spawn(resolvedCommand.command, fullArgs, {
        stdio: ["pipe", "pipe", "pipe"],
        windowsHide: true,
        env: this.childEnv()
      });
      let stdout = "";
      let stderr = "";
      const timeout = setTimeout(() => {
        child.kill();
        rejectPromise(new OfficeCliError("timeout", `officecli timed out: officecli ${args.join(" ")}`));
      }, this.options.timeoutMs ?? 120_000);

      child.stdout.on("data", (chunk: Buffer) => (stdout += chunk.toString("utf8")));
      child.stderr.on("data", (chunk: Buffer) => (stderr += chunk.toString("utf8")));
      child.on("error", (error) => {
        clearTimeout(timeout);
        rejectPromise(new OfficeCliError("spawn-error", String(error)));
      });
      child.on("close", (code) => {
        clearTimeout(timeout);
        const parsed = this.parseJson(stdout);
        if (parsed) {
          if (code === 0 && parsed.success) {
            resolvePromise(parsed);
          } else {
            rejectPromise(
              new OfficeCliError(
                parsed.error?.code ?? "failed",
                parsed.error?.error ?? parsed.message ?? `officecli exited ${code}`,
                parsed.error?.suggestion,
                stdout
              )
            );
          }
          return;
        }
        if (code === 0) {
          resolvePromise({ success: true, data: stdout.trim(), message: stdout.trim() });
        } else {
          rejectPromise(
            new OfficeCliError("non-json-exit", stderr.trim() || stdout.trim() || `exit ${code}`, undefined, stdout)
          );
        }
      });
      if (options.stdin !== undefined) {
        child.stdin.write(options.stdin);
      }
      child.stdin.end();
    });
  }

  private childEnv(): NodeJS.ProcessEnv {
    // Offline gate (§120): no self-update, no telemetry, explicit flush policy.
    return {
      ...process.env,
      OFFICECLI_NO_UPDATE: "1",
      OFFICECLI_NO_UPDATE_NOTIFIER: "1",
      OFFICECLI_RESIDENT_FLUSH: "off",
      NO_COLOR: "1"
    };
  }

  private parseJson(stdout: string): OfficeCliJson | undefined {
    const trimmed = stdout.trim();
    if (!trimmed.startsWith("{")) return undefined;
    try {
      return JSON.parse(trimmed) as OfficeCliJson;
    } catch {
      return undefined;
    }
  }

  private async resolve(): Promise<ResolvedCommand> {
    if (this.resolved) return this.resolved;
    this.resolutionPromise ??= this.resolveOnce();
    this.resolved = await this.resolutionPromise;
    return this.resolved;
  }

  private async resolveOnce(): Promise<ResolvedCommand> {
    const entry = process.env.OFFICECLI_ENTRY;
    if (entry) return { command: process.execPath, baseArgs: [resolve(entry)] };
    const bin = process.env.OFFICECLI_BIN;
    if (bin) return { command: resolve(bin), baseArgs: [] };
    if (process.platform === "win32") {
      const jsEntry = await this.resolveWindowsJsEntry();
      if (jsEntry) return { command: process.execPath, baseArgs: [jsEntry] };
      return { command: "officecli.cmd", baseArgs: [] };
    }
    return { command: "officecli", baseArgs: [] };
  }

  /** Read the officecli.cmd npm shim and extract the real JS entry path. */
  private async resolveWindowsJsEntry(): Promise<string | undefined> {
    const candidates = (process.env.PATH ?? "").split(";").filter(Boolean);
    for (const dir of candidates) {
      const cmdPath = join(dir, "officecli.cmd");
      if (!(await access(cmdPath, constants.R_OK).then(() => true, () => false))) continue;
      const content = await readFile(cmdPath, "utf8").catch(() => undefined);
      if (!content) continue;
      const match = content.match(/"([^"]+\.js)"/);
      if (match?.[1]) {
        // npm shims reference the entry relative to %dp0% (the shim's dir).
        const relative = match[1].replace(/^%dp0%[\\/]/i, "").replace(/^%~dp0[\\/]/i, "");
        const jsPath = resolve(dirname(cmdPath), relative);
        if (await access(jsPath, constants.R_OK).then(() => true, () => false)) {
          return jsPath;
        }
      }
    }
    return undefined;
  }

  async dispose(): Promise<void> {
    // Residents are tracked by the pool; nothing to do here.
  }
}
