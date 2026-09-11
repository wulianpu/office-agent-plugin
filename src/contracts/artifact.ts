/**
 * Phase 0 — Contracts: artifact plane (Design Doc §15–§27, §78–§79, §104).
 *
 * The artifact plane owns high-frequency reads. It must never require
 * Candidate/WriterLease/OfficeCLI machinery to serve a preview (P3, PERF-01/02).
 */

import type { ArtifactRef, OfficeFormat } from "./ids.js";
import type { SchedulerPriorityName } from "./scheduler.js";

/** Two-level identity (§16–§17): fast fingerprint for preview/cache, SHA-256 for write paths. */
export interface FileFingerprint {
  size: bigint;
  mtimeNs: bigint;
  fileId?: string;
}

export interface ArtifactVersionKey {
  artifactRef: ArtifactRef;
  fingerprint: FileFingerprint;
  /** SHA-256 content hash; present only on the strong-consistency path. */
  contentHash?: string;
}

/** §18: optimistic (preview) vs stable (edit/agent/commit). */
export type ReadConsistency = "optimistic" | "stable";

/** §14: frozen open intent ladder preview → open → edit. */
export type OpenIntent = "preview" | "open" | "edit";

/**
 * Core read-side identity (§19). Immutable in content, shareable,
 * reference counted via ArtifactLease, evictable, progressively enrichable (§25).
 */
export interface ArtifactContext {
  readonly artifactRef: ArtifactRef;
  readonly version: ArtifactVersionKey;
  readonly format: OfficeFormat;
  readonly consistency: ReadConsistency;
  readonly rendererVersion: string;
  readonly lastAccessAt: number;

  /** Lazy enrichment payload owned by the FormatRuntime (§25). Content identity never mutates. */
  readonly enrichment: Map<string, unknown>;
}

export interface ArtifactLease {
  readonly leaseId: string;
  readonly context: ArtifactContext;
  /** Consumer identity used for debug introspection (§22) and consumer-aware cancellation (§24). */
  readonly consumer: string;
  release(): void;
}

export interface ArtifactBuildInput {
  artifactRef: ArtifactRef;
  format: OfficeFormat;
  consistency: ReadConsistency;
  priority: SchedulerPriorityName;
  consumer: string;
  signal?: AbortSignal;
}

/** §21: every context acquisition goes through the registry — no ad-hoc parsing. */
export interface ArtifactRegistry {
  acquire(input: ArtifactBuildInput): Promise<ArtifactLease>;
  /** Resolve the currently cached context for a version key without building. */
  peek(key: ArtifactVersionKey): ArtifactContext | undefined;
  /** Debug view: who still holds a context (§22). */
  holders(artifactRef: ArtifactRef): ReadonlyArray<{ leaseId: string; consumer: string }>;
  /** Evict completed contexts whose refcount reached zero beyond the retention budget. */
  trim(level: MemoryTrimLevel): Promise<void>;
}

/** Per-format process-wide runtime sharing fonts/workers/parsers (§26). */
export interface FormatRuntime {
  readonly format: OfficeFormat;
  initialize(): Promise<void>;
  createArtifactContext(input: ArtifactBuildInput): Promise<ArtifactContext>;
  trimMemory(level: MemoryTrimLevel): Promise<void>;
  dispose(): Promise<void>;
}

export type MemoryTrimLevel = "light" | "moderate" | "aggressive";

/** §79: hot path probe — stat + format + ZIP central directory only. Never full hash. */
export interface FastArtifactProbeResult {
  artifactRef: ArtifactRef;
  format: OfficeFormat;
  fingerprint: FileFingerprint;
  entryCount: number;
  /** True when [Content_Types].xml is present and ZIP structure is sane. */
  packageSane: boolean;
  probedAt: number;
}

/** §78: strong scan, performed once and shared by hash/verification/diff/revision metadata. */
export interface PackageManifestEntry {
  name: string;
  compressedSize: number;
  size: number;
  crc32: number;
  method: number;
}

export interface PackageManifest {
  entries: PackageManifestEntry[];
}

export interface IntegrityResult {
  ok: boolean;
  /** Entry names whose local data fails CRC/structure checks. */
  corruptEntries: string[];
}

export interface ArtifactScanResult {
  contentHash: string;
  size: bigint;
  packageManifest: PackageManifest;
  integrity: IntegrityResult;
  scannedAt: number;
}

/** §104: preview cache key invalidates on renderer/font/profile changes. */
export interface PreviewCacheKey {
  contentHash: string;
  rendererVersion: string;
  fontEnvironmentId: string;
  previewProfile: string;
}

/** §113–§114: binary IPC payload ownership. */
export interface BinaryPayload {
  buffer: ArrayBuffer;
  ownership: "transfer" | "copy";
}
