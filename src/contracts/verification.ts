/**
 * Phase 0 — Contracts: verification plane (Design Doc §80–§84).
 */

import type { CandidateId } from "./ids.js";

export type VerificationConfidence =
  | "unverified"
  | "structural"
  | "engine"
  | "visual"
  | "consumer-certified";

/** §81: layered checks L0..L7. */
export type VerificationLayer =
  | "L0-artifact-integrity"
  | "L1-ooxml-structural"
  | "L2-package-relationships"
  | "L3-officecli-issues"
  | "L4-mutation-aware-diff"
  | "L5-changed-scope-render"
  | "L6-genoffice-preview"
  | "L7-host-certification";

export interface CheckResult {
  layer: VerificationLayer;
  status: "pass" | "warn" | "fail" | "skipped";
  /** Issues found; empty on pass. */
  issues: VerificationIssue[];
  durationMs: number;
}

export interface VerificationIssue {
  severity: "info" | "warn" | "error";
  code: string;
  message: string;
  /** Package part or element path the issue refers to. */
  target?: string;
}

/** §82: verification is always bound to a candidate contentHash (INV-08). */
export interface VerificationReport {
  candidateId: CandidateId;
  contentHash: string;
  structural: CheckResult;
  package: CheckResult;
  semantic: CheckResult;
  visual: CheckResult;
  confidence: VerificationConfidence;
  verifiedAt: number;
}

/** §80: package diff risk classification. */
export type PartDiffClass =
  | "UNCHANGED"
  | "NORMALIZATION_ONLY"
  | "EXPECTED"
  | "UNEXPECTED_LOW_RISK"
  | "UNEXPECTED_HIGH_RISK";

export interface PackageDiffEntry {
  name: string;
  change: "added" | "removed" | "modified" | "unchanged";
  classification: PartDiffClass;
}

export interface PackageDiffResult {
  entries: PackageDiffEntry[];
  summary: Record<PartDiffClass, number>;
  /** Parts whose relationship graph changed (relationships, media refs). */
  relationshipChanges: string[];
  highRiskParts: string[];
}
