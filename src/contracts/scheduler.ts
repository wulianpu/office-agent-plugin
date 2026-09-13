/**
 * Phase 0 — Contracts: resource & scheduling plane (Design Doc §105–§112, §89–§101).
 */

import type { MemoryTrimLevel } from "./artifact.js";

/** §110: scheduler priority ladder — foreground always outranks background (PERF-10). */
export const SCHEDULER_PRIORITIES = [
  "INTERACTIVE",
  "VISIBLE_PREVIEW",
  "EDIT_PROMOTION",
  "AGENT_FOREGROUND",
  "VERIFICATION",
  "PREFETCH",
  "BACKGROUND_INDEX",
  "CACHE_BUILD"
] as const;

export type SchedulerPriorityName = (typeof SCHEDULER_PRIORITIES)[number];

/** §106: budgets are split per resource class, never just the JS heap. */
export type ResourceClass =
  | "memory"
  | "cpu"
  | "io"
  | "render"
  | "native-process"
  | "xlsx-sidecar"
  | "disk-cache";

export interface ResourceRequest {
  classes: Partial<Record<ResourceClass, number>>;
  /** Objects protected from pressure eviction: unsaved human state, candidates, commit journal (§107). */
  protected?: boolean;
  /** Payload byte size for memory accounting against the byte budget (§101). */
  bytes?: number;
  label: string;
}

export interface ResourcePressureSnapshot {
  level: ResourcePressure;
  utilization: Partial<Record<ResourceClass, number>>;
  timestamp: number;
}

export type ResourcePressure = "normal" | "elevated" | "high" | "critical";

export interface ResourceLease {
  readonly leaseId: string;
  readonly request: ResourceRequest;
  release(): void;
}

/** §105: ResourceGovernor owns global budgets and is P0 infrastructure. */
export interface ResourceGovernor {
  acquire(request: ResourceRequest): Promise<ResourceLease>;
  tryAcquire(request: ResourceRequest): ResourceLease | undefined;
  getPressure(): ResourcePressureSnapshot;
  /** Register a trim target invoked under memory pressure (eviction ladder order §107). */
  registerTrimTarget(target: ResourceTrimTarget): void;
  onPressureChange(listener: (snapshot: ResourcePressureSnapshot) => void): () => void;
}

export interface ResourceTrimTarget {
  /** Lower rank = trimmed earlier (§107 ladder). */
  rank: number;
  label: string;
  trim(level: MemoryTrimLevel): Promise<void>;
}

export type Residency = "hot" | "warm" | "cold" | "evicted";

/** §99: caches are typed separately; each has its own lifecycle and byte budget. */
export type CacheKind =
  | "ArtifactCache"
  | "VisualPreviewCache"
  | "DecodedAssetCache"
  | "EditorRuntimeCache";

export interface CacheAdmissionPolicy {
  /** Skip caching when encoded size exceeds this share of the budget (§100). */
  maxItemShareOfBudget: number;
}

export interface CacheStats {
  kind: CacheKind;
  hits: number;
  misses: number;
  hitRate(): number;
}
