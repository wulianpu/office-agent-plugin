/**
 * Phase 0 — Contracts: preview plane (Design Doc §10–§12, §18).
 */

import type { ArtifactRef } from "./ids.js";
import type { PreviewCacheKey } from "./artifact.js";

export type PreviewPriority = "visible" | "prefetch" | "background";

export interface PreviewScope {
  /** e.g. slide index, sheet name + range, block anchor. */
  location?: unknown;
  maxEntries?: number;
}

/** §11: preview requests create no DocumentSession/lease/candidate/OfficeCLI (PERF-01/02/03). */
export interface PreviewRequest {
  requestId: string;
  artifactRef: ArtifactRef;
  priority: PreviewPriority;
  scope?: PreviewScope;
  /**
   * P0-7: default previews stay at ZIP/index depth (first useful content,
   * §28-30); `visual` upgrades to the full engine model + SVG slides.
   */
  visual?: boolean;
}

/** Headless render model consumed by preview surfaces. */
export interface PreviewModel {
  format: "docx" | "xlsx" | "pptx";
  /** Structured outline: blocks / sheet grid window / slide texts. */
  outline: PreviewOutline;
  cacheKey: PreviewCacheKey;
  fingerprintAtRender: string;
  /** PPTX visual rendering: standalone SVG per slide (headless, §P7 data plane). */
  svgSlides?: string[];
}

export type PreviewOutline =
  | { kind: "docx"; blocks: Array<{ index: number; style?: string; text: string }> }
  | { kind: "xlsx"; sheets: Array<{ name: string; rowCount: number; window: string[][] }> }
  | { kind: "pptx"; slides: Array<{ index: number; shapes: Array<{ name?: string; text?: string }> }> };

export interface PreviewResult {
  requestId: string;
  model: PreviewModel;
  /** True when the fingerprint changed during render and the result was recomputed (§18). */
  retried: boolean;
  elapsedMs: number;
}
