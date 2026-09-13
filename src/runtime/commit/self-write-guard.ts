/**
 * SelfWriteGuard (§73): the runtime's own commits produce filesystem watcher
 * events. Guards match those events by expected hash and expiry so a commit
 * never cascades into "external conflict → reload".
 */

import type { SelfWriteGuard } from "../../contracts/revision.js";
import { sha256File } from "../../support/fsx.js";

export class SelfWriteGuardRegistry {
  private readonly guards = new Map<string, SelfWriteGuard & { sourcePath: string }>();

  register(guard: SelfWriteGuard & { sourcePath: string }): void {
    this.guards.set(guard.commitId, guard);
    this.sweep();
  }

  /**
   * When a watcher event fires for a path: if an active guard exists and the
   * file now hashes to the guard's expected value, the event is self-originated.
   * Comparison is long-form on both sides (watcher events arrive long-form).
   */
  async isSelfWrite(sourcePath: string): Promise<boolean> {
    this.sweep();
    const canonical = longFormPath(sourcePath);
    for (const guard of this.guards.values()) {
      if (longFormPath(guard.sourcePath) !== canonical) continue;
      if (Date.now() > guard.expiresAt) continue;
      const hash = await sha256File(canonical).catch(() => undefined);
      if (hash === guard.expectedHash) return true;
    }
    return false;
  }

  private sweep(): void {
    const now = Date.now();
    for (const [id, guard] of this.guards) {
      if (now > guard.expiresAt + 60_000) this.guards.delete(id);
    }
  }

  get size(): number {
    return this.guards.size;
  }
}

/**
 * SourceWatcher: filesystem watcher with self-write suppression and external
 * mutation detection (conflict → session.lifecycle = conflict, §77 truth from
 * filesystem hashes).
 */

import { realpathSync, watch, type FSWatcher } from "node:fs";
import { basename, dirname, join, resolve } from "node:path";

export interface SourceMutationEvent {
  sourcePath: string;
  kind: "self-write" | "external";
}

/**
 * Canonical long-form path (dir via realpath + basename). Windows 8.3 short
 * paths (e.g. RUNNER~1 from %TEMP%) crash libuv's fs-event watcher: the OS
 * reports long-name files whose prefix no longer matches the short watched
 * directory — a C-level assertion that ABORTS the process. Every path that
 * reaches watch()/watcher callbacks is normalized through here so both sides
 * always compare in the same form.
 */
export function longFormPath(path: string): string {
  const absolute = resolve(path);
  try {
    return join(realpathSync(dirname(absolute)), basename(absolute));
  } catch {
    return absolute;
  }
}

export class SourceWatcher {
  /**
   * Directory watchers with ref-counted path sets (P0-3): one FSWatcher per
   * directory covers every watched file inside it; watching and unwatching
   * are per-file and independent.
   */
  private readonly dirs = new Map<string, { watcher: FSWatcher; paths: Map<string, number> }>();
  private readonly listeners = new Set<(event: SourceMutationEvent) => void>();
  private pending = new Map<string, NodeJS.Timeout>();

  constructor(private readonly selfWrites: SelfWriteGuardRegistry) {}

  watchFile(sourcePath: string): void {
    const canonical = longFormPath(sourcePath);
    const dir = dirname(canonical);
    let entry = this.dirs.get(dir);
    if (!entry) {
      try {
        const watcher = watch(dir, (event, filename) => {
          if (event !== "change" && event !== "rename") return;
          const name = filename ? String(filename) : "";
          if (!name) return;
          const changed = resolve(dir, name);
          if ((this.dirs.get(dir)?.paths.get(changed) ?? 0) > 0) this.schedule(changed);
        });
        entry = { watcher, paths: new Map() };
        this.dirs.set(dir, entry);
      } catch {
        // Watchers are best-effort; conflict detection falls back to hash checks.
        return;
      }
    }
    entry.paths.set(canonical, (entry.paths.get(canonical) ?? 0) + 1);
  }

  unwatchFile(sourcePath: string): void {
    const canonical = longFormPath(sourcePath);
    const dir = dirname(canonical);
    const entry = this.dirs.get(dir);
    if (!entry) return;
    const count = entry.paths.get(canonical) ?? 0;
    if (count <= 1) {
      entry.paths.delete(canonical);
      const timer = this.pending.get(canonical);
      if (timer) {
        clearTimeout(timer);
        this.pending.delete(canonical);
      }
    } else {
      entry.paths.set(canonical, count - 1);
    }
    if (entry.paths.size === 0) {
      entry.watcher.close();
      this.dirs.delete(dir);
    }
  }

  private schedule(sourcePath: string): void {
    const existing = this.pending.get(sourcePath);
    if (existing) clearTimeout(existing);
    const timer = setTimeout(() => {
      this.pending.delete(sourcePath);
      void this.dispatch(sourcePath);
    }, 250);
    this.pending.set(sourcePath, timer);
  }

  private async dispatch(sourcePath: string): Promise<void> {
    sourcePath = resolve(sourcePath);
    if (await this.selfWrites.isSelfWrite(sourcePath)) {
      for (const listener of this.listeners) listener({ sourcePath, kind: "self-write" });
      return;
    }
    for (const listener of this.listeners) listener({ sourcePath, kind: "external" });
  }

  onMutation(listener: (event: SourceMutationEvent) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  /** Watched-file count (sum of refcounts) — test/telemetry introspection. */
  watchedFileCount(): number {
    let total = 0;
    for (const counts of this.dirs.values()) {
      for (const n of counts.paths.values()) total += n;
    }
    return total;
  }

  /** Distinct watched directories — OS watcher footprint. */
  directoryCount(): number {
    return this.dirs.size;
  }

  dispose(): void {
    for (const entry of this.dirs.values()) entry.watcher.close();
    this.dirs.clear();
    for (const timer of this.pending.values()) clearTimeout(timer);
    this.pending.clear();
  }
}
