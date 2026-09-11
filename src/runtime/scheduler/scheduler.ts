/**
 * Scheduler (§109–§112): Session Actors dispatch heavy jobs here; the ladder
 * guarantees interactive work outranks background work (PERF-08, PERF-10) and
 * queues apply backpressure instead of growing without bound.
 */

import {
  SCHEDULER_PRIORITIES,
  type SchedulerPriorityName
} from "../../contracts/scheduler.js";
import type { AsyncWorkIdentity } from "../../contracts/document.js";
import { isAsyncWorkStale } from "../../contracts/document.js";
import type { ResourceGovernor } from "../resources/resource-governor.js";

export interface SchedulerJob<T> {
  label: string;
  priority: SchedulerPriorityName;
  /**
   * P1e (§105-§106): resource classes this job consumes. When a governor is
   * attached, the scheduler acquires/releases real leases around execution —
   * budgets actually gate work instead of being advisory accounting.
   */
  resources?: {
    cpu?: number;
    io?: number;
    render?: number;
    nativeProcess?: number;
    estimatedBytes?: number;
  };
  /** Freshness identity checked before start and before result delivery (§43). */
  identity?: AsyncWorkIdentity;
  current?: () => { sessionEpoch?: number; fencingToken?: bigint; candidateId?: string } | undefined;
  run: (signal: AbortSignal) => Promise<T>;
  /** Stale jobs are dropped silently when superseded (backpressure §112). */
  onStale?: () => void;
}

export interface ScheduledHandle<T> {
  promise: Promise<T>;
  cancel(): void;
}

function resourcesToRequest(
  resources: NonNullable<SchedulerJob<unknown>["resources"]>,
  label: string
): { classes: Record<string, number>; bytes?: number; label: string } {
  const classes: Record<string, number> = {};
  if (resources.cpu) classes.cpu = resources.cpu;
  if (resources.io) classes.io = resources.io;
  if (resources.render) classes.render = resources.render;
  if (resources.nativeProcess) classes["native-process"] = resources.nativeProcess;
  return { classes, bytes: resources.estimatedBytes, label: `job:${label}` };
}

const PRIORITY_INDEX = new Map<SchedulerPriorityName, number>(
  SCHEDULER_PRIORITIES.map((name, index) => [name, index])
);

interface QueuedJob {
  id: number;
  priorityIndex: number;
  enqueueAt: number;
  job: SchedulerJob<unknown>;
  abort: AbortController;
  resolve: (value: unknown) => void;
  reject: (error: unknown) => void;
}

export interface SchedulerOptions {
  maxConcurrent?: number;
  maxQueuePerPriority?: number;
  /** P1e: when present, jobs carrying `resources` acquire real leases. */
  governor?: ResourceGovernor;
}

export class Scheduler {
  private queue: QueuedJob[] = [];
  private running = 0;
  private nextId = 1;
  private readonly maxConcurrent: number;
  private readonly maxQueuePerPriority: number;
  readonly stats = { started: 0, completed: 0, cancelled: 0, staleDropped: 0, backpressureDrops: 0 };

  private readonly governor?: ResourceGovernor;

  constructor(options: SchedulerOptions = {}) {
    this.maxConcurrent = options.maxConcurrent ?? 8;
    this.maxQueuePerPriority = options.maxQueuePerPriority ?? 256;
    this.governor = options.governor;
    // Resources freed outside this scheduler (e.g. sidecar leases) must
    // re-kick the pump so blocked jobs get their chance — on EVERY release,
    // not only pressure-level transitions.
    this.governor?.onRelease(() => this.pump());
  }

  submit<T>(job: SchedulerJob<T>): ScheduledHandle<T> {
    const priorityIndex = PRIORITY_INDEX.get(job.priority) ?? SCHEDULER_PRIORITIES.length;
    const queued: QueuedJob = {
      id: this.nextId++,
      priorityIndex,
      enqueueAt: Date.now(),
      job: job as SchedulerJob<unknown>,
      abort: new AbortController(),
      resolve: undefined as never,
      reject: undefined as never
    };
    let resolve!: (value: unknown) => void;
    let reject!: (error: unknown) => void;
    const promise = new Promise<T>((res, rej) => {
      resolve = res as (value: unknown) => void;
      reject = rej;
    });
    queued.resolve = resolve;
    queued.reject = reject;

    // Backpressure: cap each lane; drop the oldest queued job of the same lane.
    const sameLane = this.queue.filter((q) => q.priorityIndex === priorityIndex);
    if (sameLane.length >= this.maxQueuePerPriority) {
      const oldest = sameLane[0]!;
      this.queue = this.queue.filter((q) => q !== oldest);
      this.stats.backpressureDrops++;
      oldest.reject(new Error(`scheduler backpressure: dropped '${oldest.job.label}'`));
    }

    this.queue.push(queued);
    this.queue.sort((a, b) => a.priorityIndex - b.priorityIndex || a.enqueueAt - b.enqueueAt);
    this.pump();
    return {
      promise,
      cancel: () => {
        const idx = this.queue.indexOf(queued);
        if (idx >= 0) {
          this.queue.splice(idx, 1);
          this.stats.cancelled++;
          queued.reject(new DOMException("cancelled", "AbortError"));
          return;
        }
        queued.abort.abort();
      }
    };
  }

  /** Cancel queued (not started) jobs matching a predicate — used when revisions supersede preview work. */
  cancelWhere(predicate: (job: SchedulerJob<unknown>) => boolean): number {
    let removed = 0;
    this.queue = this.queue.filter((q) => {
      if (predicate(q.job)) {
        removed++;
        this.stats.cancelled++;
        q.reject(new DOMException("cancelled", "AbortError"));
        return false;
      }
      return true;
    });
    return removed;
  }

  get queueDepth(): number {
    return this.queue.length;
  }

  get runningCount(): number {
    return this.running;
  }

  /**
   * P1-high-C: governor-aware dispatch. A job that only WAITS for resources
   * must not occupy a running slot (priority inversion: 8 blocked background
   * jobs would starve INTERACTIVE). The pump scans in priority order for the
   * first job whose resources are acquirable NOW; blocked jobs stay queued.
   */
  private pumping = false;

  private pump(): void {
    // Re-entrancy guard: tryAcquire can synchronously fire pressure-change
    // listeners that call pump() again — the nested call must not splice
    // entries the outer scan is about to pick. The outer loop (or a finally)
    // re-drives once done.
    if (this.pumping) return;
    this.pumping = true;
    try {
      this.pumpInner();
    } finally {
      this.pumping = false;
    }
  }

  private pumpInner(): void {
    while (this.running < this.maxConcurrent && this.queue.length > 0) {
      let pickedIndex = -1;
      let pickedLease: ReturnType<ResourceGovernor["tryAcquire"]> = undefined;
      for (let i = 0; i < this.queue.length; i++) {
        const candidate = this.queue[i]!;
        if (candidate.abort.signal.aborted) {
          this.queue.splice(i, 1);
          this.stats.cancelled++;
          candidate.reject(new DOMException("cancelled", "AbortError"));
          i--;
          continue;
        }
        if (!candidate.job.resources || !this.governor) {
          pickedIndex = i;
          break;
        }
        const lease = this.governor.tryAcquire(resourcesToRequest(candidate.job.resources, candidate.job.label));
        if (lease) {
          pickedIndex = i;
          pickedLease = lease;
          break;
        }
        // Blocked at this priority — keep scanning lower priorities.
      }
      if (pickedIndex < 0) return; // every queued job is resource-blocked
      const next = this.queue.splice(pickedIndex, 1)[0]!;
      this.running++;
      this.stats.started++;
      void this.execute(next, pickedLease ?? undefined);
    }
  }

  private async execute(queued: QueuedJob, preAcquiredLease?: NonNullable<ReturnType<ResourceGovernor["tryAcquire"]>>): Promise<void> {
    const { job } = queued;
    const staleBefore = this.isStale(job);
    if (staleBefore) {
      // P0-4: the pump pre-acquired the resource lease for this job — bailing
      // before try/finally must still release it, or io/render slots drain.
      preAcquiredLease?.release();
      this.stats.staleDropped++;
      this.stats.completed++;
      this.running--;
      job.onStale?.();
      queued.resolve(undefined);
      this.pump();
      return;
    }
    // P1e/P1-high-C: the pump pre-acquired resource leases; a job WITHOUT
    // resources (or when no governor is attached) runs lease-free.
    const lease = preAcquiredLease;
    try {
      const result = await job.run(queued.abort.signal);
      if (this.isStale(job)) {
        this.stats.staleDropped++;
        job.onStale?.();
        queued.resolve(undefined);
      } else {
        queued.resolve(result);
      }
    } catch (error) {
      queued.reject(error);
    } finally {
      lease?.release();
      this.stats.completed++;
      this.running--;
      this.pump();
    }
  }

  private isStale(job: SchedulerJob<unknown>): boolean {
    if (!job.identity || !job.current) return false;
    const current = job.current();
    if (!current) return false;
    return isAsyncWorkStale(job.identity, current);
  }
}
