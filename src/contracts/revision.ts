/**
 * Phase 0 — Contracts: revision plane (Design Doc §36, §71–§77).
 */

import type { ArtifactRef, RevisionId, SessionId } from "./ids.js";

/**
 * §36: committed revisions are the MVCC read anchors. Human Ctrl+S and Agent
 * Accept both land here — hence "committed", not "accepted".
 */
export interface CommittedRevision {
  revisionId: RevisionId;
  sessionId: SessionId;
  sequence: number;
  artifactRef: ArtifactRef;
  contentHash: string;
  origin: RevisionOrigin;
  createdAt: number;
  /**
   * Exact commit identity (schema v3): set for every revision landed through
   * the atomic committer or forward-recovery. UNIQUE in SQLite — a commit can
   * land at most one revision regardless of content hash. Null only for
   * legacy rows and journal-less external revisions.
   */
  commitId?: string;
}

export type RevisionOrigin = "human" | "agent" | "external";

/** §74–§76: every source replacement goes through the atomic committer with a journal. */
export interface CommitRequest {
  commitId: string;
  sessionId: SessionId;
  candidateId: string;
  /** Source artifact the candidate replaces. */
  artifactRef: ArtifactRef;
  /** Candidate artifact that becomes the new committed content. */
  candidateArtifactRef: ArtifactRef;
  /** Hash the source must still have at replace time (INV-10). */
  expectedSourceHash: string;
  candidateHash: string;
  origin: RevisionOrigin;
}

export interface CommitResult {
  commitId: string;
  newRevision: CommittedRevision;
  /** How the journal resolved for this commit (used by tests/recovery introspection). */
  journalOutcome: "finalized" | "rolled-back";
}

export type CommitPhase =
  | "prepared"
  | "temp-ready"
  | "source-replaced"
  | "finalized"
  | "aborted";

/** §76: journal row shape persisted in SQLite. */
export interface CommitJournalRecord {
  commitId: string;
  sessionId: SessionId;
  candidateId: string;
  sourcePath: string;
  tempPath: string;
  sourceHashBefore: string;
  candidateHash: string;
  phase: CommitPhase;
  origin: RevisionOrigin;
  createdAt: number;
  updatedAt: number;
}

/** §73: identifies self-originated watcher events after a commit. */
export interface SelfWriteGuard {
  commitId: string;
  expectedHash: string;
  expiresAt: number;
}

export type RecoveryOutcome =
  | { commitId: string; resolution: "finalized"; contentHash: string }
  | { commitId: string; resolution: "rolled-back"; reason: "source-unchanged" | "temp-only" | "aborted" }
  | { commitId: string; resolution: "conflict"; reason: string };
