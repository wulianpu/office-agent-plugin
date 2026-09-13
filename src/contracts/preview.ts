/**
 * Phase 0 — Contracts: preview plane (Design Doc §10–§12, §18).
 */

import type { ArtifactRef } from "./ids.js";
import type { PreviewCacheKey } from "./artifact.js";

export type PreviewPriority = "visible" | "prefetch" | "background";

/** Typed scope location (round 10): per-format anchor into the document. */
export interface PreviewScopeLocation {
  /** XLSX sheet name to window into. */
  sheet?: string;
  /** PPTX 1-based slide number the window starts from. */
  slide?: number;
  /** DOCX 0-based block anchor the window starts from. */
  block?: number;
  /** XLSX bounded range, 1-based inclusive. */
  range?: { fromRow: number; toRow: number; fromCol: number; toCol: number };
}

export interface PreviewScope {
  location?: PreviewScopeLocation;
  /** Window size cap for entries (rows/blocks/slides). */
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
   * §28–30); `visual` upgrades to the full engine model + SVG slides.
   */
  visual?: boolean;
  /**
   * Round 10: consumer cancellation. Aborting dequeues queued scheduler
   * work before it reaches the sidecar/engine; in-flight renders finish.
   */
  signal?: AbortSignal;
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
  | {
      kind: "xlsx";
      sheets: Array<{
        name: string;
        rowCount: number;
        /** Round 10: false when rowCount is a scanned lower bound (fallback
         *  without a declared dimension), true when exact (sidecar metadata
         *  or worksheet <dimension>). */
        rowCountExact?: boolean;
        window: string[][];
      }>;
    }
  | { kind: "pptx"; slides: Array<{ index: number; shapes: Array<{ name?: string; text?: string }> }> };

export interface PreviewResult {
  requestId: string;
  model: PreviewModel;
  /** True when the fingerprint changed during render and the result was recomputed (§18). */
  retried: boolean;
  elapsedMs: number;
}
