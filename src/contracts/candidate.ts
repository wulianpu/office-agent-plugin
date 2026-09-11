/**
 * Phase 0 — Contracts: candidate plane (Design Doc §35, §37, §66–§72, §84).
 */

import type { ArtifactRef, CandidateId, RevisionId, SessionId } from "./ids.js";
import type { VerificationReport } from "./verification.js";

/** §35: candidate state machine, kept separate from session/editor states. */
export type CandidateState =
  | "preparing"
  | "mutating"
  | "flushing"
  | "verifying"
  | "ready"
  | "human-amended"
  | "committing"
  | "failed";

/** §37: agent (and human-amended) proposals never touch committed artifacts (P12). */
export interface CandidateRevision {
  candidateId: CandidateId;
  sessionId: SessionId;
  baseRevisionId: RevisionId;
  baseHash: string;
  artifactRef: ArtifactRef;
  currentHash?: string;
  state: CandidateState;
  createdBy: "agent" | "human";
  verification?: VerificationReport;
  failureReason?: string;
  createdAt: number;
  updatedAt: number;
}

/** §84: what a mutation touched — drives changed-scope-first verification (§83). */
export interface ChangeImpact {
  scope: "document" | "slide" | "sheet" | "range" | "block";
  targets?: string[];
  confidence: "exact" | "conservative" | "unknown";
}

export const CANDIDATE_TRANSITIONS: Readonly<Record<CandidateState, readonly CandidateState[]>> = {
  preparing: ["mutating", "failed"],
  mutating: ["flushing", "failed"],
  flushing: ["verifying", "failed"],
  verifying: ["ready", "failed"],
  ready: ["human-amended", "committing", "failed"],
  "human-amended": ["verifying", "failed"],
  committing: ["failed"],
  failed: []
};

export function canTransitionCandidate(from: CandidateState, to: CandidateState): boolean {
  return CANDIDATE_TRANSITIONS[from].includes(to);
}
