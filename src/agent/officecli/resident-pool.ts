/**
 * ResidentPool (§65): OfficeCLI residents are cacheable runtime resources —
 * never required state of a DocumentSession. The governor bounds concurrency
 * and idle residents are evicted after a TTL.
 */

import type { ResourceGovernor } from "../../contracts/scheduler.js";
import type { ResourceLease } from "../../contracts/scheduler.js";
import type { OfficeCliAdapter } from "./officecli-adapter.js";

interface ResidentEntry {
  file: string;
  lease: ResourceLease;
  lastUsedAt: number;
}

export class ResidentPool {
  private readonly residents = new Map<string, ResidentEntry>();
  /**
   * P1-high (#6): same-file opening single-flight. The pool is the resource
   * ownership boundary — it must never depend on callers being serialized.
   * Without this, two concurrent acquire(file) calls each took a governor
   * lease and each ran adapter.open(); the second Map.set overwrote the
   * first entry, orphaning its ResourceLease (unreleasable until dispose).
   */
  private readonly openings = new Map<string, Promise<ResidentEntry>>();
  private sweeper?: NodeJS.Timeout;
  private sweeping = false;

  constructor(
    private readonly adapter: OfficeCliAdapter,
    private readonly governor: ResourceGovernor,
    private readonly options: { idleTtlMs?: number; sweepIntervalMs?: number } = {}
  ) {}

  private get idleTtlMs(): number {
    return this.options.idleTtlMs ?? 120_000;
  }

  /** Ensure a live resident for the file and return a usage handle. */
  async acquire(file: string): Promise<ResidentHandle> {
    const existing = this.residents.get(file);
    if (existing) {
      existing.lastUsedAt = Date.now();
      return { file, pool: this, closeOnRelease: false };
    }
    let opening = this.openings.get(file);
    if (!opening) {
      opening = this.openResident(file).finally(() => this.openings.delete(file));
      this.openings.set(file, opening);
    }
    await opening;
    return { file, pool: this, closeOnRelease: false };
  }

  /** One governor lease + one adapter.open per file, shared by all joiners;
   *  on failure the lease is released exactly once and every joiner sees the
   *  same error with the pool back at baseline. */
  private async openResident(file: string): Promise<ResidentEntry> {
    const lease = await this.governor.acquire({
      classes: { "native-process": 1, memory: 1 },
      bytes: 32 * 1024 * 1024,
      label: `officecli-resident:${file}`,
      protected: false
    });
    try {
      await this.adapter.open(file);
    } catch (error) {
      lease.release();
      throw error;
    }
    const entry: ResidentEntry = { file, lease, lastUsedAt: Date.now() };
    this.residents.set(file, entry);
    this.ensureSweeper();
    return entry;
  }

  /**
   * Release a usage handle. The resident stays warm unless the caller asks
   * for a handoff close (§64: close = flush + release).
   */
  async release(handle: ResidentHandle): Promise<void> {
    if (handle.closeOnRelease) {
      await this.evict(handle.file);
      return;
    }
    const entry = this.residents.get(handle.file);
    if (entry) entry.lastUsedAt = Date.now();
  }

  /** Flush + release a resident immediately (INV-07 before writer handoff). */
  async evict(file: string): Promise<void> {
    const entry = this.residents.get(file);
    if (!entry) return;
    this.residents.delete(file);
    try {
      await this.adapter.close(file);
    } catch (error) {
      // One retry: the resident may be mid-idle-flush; a silent leak would
      // hold Windows file locks forever (§75).
      try {
        await new Promise((resolve) => setTimeout(resolve, 300));
        await this.adapter.close(file);
      } catch (retryError) {
        entry.lease.release();
        throw retryError ?? error;
      }
    } finally {
      if (this.residents.size === 0) this.stopSweeper();
    }
    entry.lease.release();
  }

  /** Idle eviction for the memory-pressure ladder (rank: idle-residents). */
  async trimIdle(): Promise<number> {
    const now = Date.now();
    let evicted = 0;
    for (const [file, entry] of [...this.residents]) {
      if (now - entry.lastUsedAt > this.idleTtlMs) {
        await this.evict(file);
        evicted++;
      }
    }
    return evicted;
  }

  size(): number {
    return this.residents.size;
  }

  private ensureSweeper(): void {
    if (this.sweeper) return;
    this.sweeper = setInterval(() => {
      void this.trimIdle().catch(() => undefined);
    }, this.options.sweepIntervalMs ?? 30_000);
    this.sweeper.unref?.();
  }

  private stopSweeper(): void {
    if (this.sweeper) {
      clearInterval(this.sweeper);
      this.sweeper = undefined;
    }
  }

  async dispose(): Promise<void> {
    this.stopSweeper();
    for (const file of [...this.residents.keys()]) {
      await this.evict(file);
    }
  }
}

export interface ResidentHandle {
  file: string;
  pool: ResidentPool;
  closeOnRelease: boolean;
}
