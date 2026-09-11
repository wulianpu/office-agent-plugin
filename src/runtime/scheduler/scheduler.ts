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

export interface SchedulerJob<T> {
  label: string;
  priority: SchedulerPriorityName;
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
}

export class Scheduler {
  private queue: QueuedJob[] = [];
  private running = 0;
  private nextId = 1;
  private readonly maxConcurrent: number;
  private readonly maxQueuePerPriority: number;
  readonly stats = { started: 0, completed: 0, cancelled: 0, staleDropped: 0, backpressureDrops: 0 };

  constructor(options: SchedulerOptions = {}) {
    this.maxConcurrent = options.maxConcurrent ?? 8;
    this.maxQueuePerPriority = options.maxQueuePerPriority ?? 256;
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

  private pump(): void {
    while (this.running < this.maxConcurrent && this.queue.length > 0) {
      const next = this.queue.shift()!;
      if (next.abort.signal.aborted) continue;
      this.running++;
      this.stats.started++;
      void this.execute(next);
    }
  }

  private async execute(queued: QueuedJob): Promise<void> {
    const { job } = queued;
    const staleBefore = this.isStale(job);
    if (staleBefore) {
      this.stats.staleDropped++;
      this.stats.completed++;
      this.running--;
      job.onStale?.();
      queued.resolve(undefined);
      this.pump();
      return;
    }
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
