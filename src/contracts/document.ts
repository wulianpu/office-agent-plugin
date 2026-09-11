/**
 * Phase 0 — Contracts: document plane (Design Doc §32–§43, §51–§54).
 */

import type {
  ArtifactRef,
  CandidateId,
  DocumentBackend,
  DocumentId,
  OfficeFormat,
  RevisionId,
  SessionEpoch,
  SessionId
} from "./ids.js";
import type { CommittedRevision } from "./revision.js";
import type { CandidateRevision } from "./candidate.js";
import type { WriterLease } from "./lease.js";
import type { EditorBinding } from "./editor.js";

/** §33: session lifecycle kept separate from editor/candidate states. */
export type SessionLifecycle =
  | "opening"
  | "ready"
  | "conflict"
  | "recovery-required"
  | "closing"
  | "closed";

export interface DocumentSession {
  readonly sessionId: SessionId;
  readonly documentId: DocumentId;
  readonly format: OfficeFormat;
  readonly backend: DocumentBackend;
  lifecycle: SessionLifecycle;
  /** Bumped on recovery/rebind so in-flight results go stale (§43). */
  sessionEpoch: SessionEpoch;
  committedRevision: CommittedRevision;
  candidate?: CandidateRevision;
  writerLease?: WriterLease;
  editor?: EditorBinding;
  readonly artifactRef: ArtifactRef;
  readonly createdAt: number;
}

/** §52: write readiness prepared after formal open, without acquiring a lease. */
export type WriteReadiness = "cold" | "probing" | "ready" | "blocked";

export interface WriteReadinessState {
  readiness: WriteReadiness;
  strongHash?: string;
  /** Human-readable reason when blocked (e.g. file locked, permission denied). */
  blockedReason?: string;
  probedAt?: number;
}

/** §54: view position preserved across preview/open→edit promotion. */
export interface ViewBookmark {
  location: unknown;
  zoom?: number;
}

/** §43: staleness identity attached to every dispatched async job. */
export interface AsyncWorkIdentity {
  sessionId: SessionId;
  sessionEpoch: SessionEpoch;
  artifactRef: ArtifactRef;
  candidateId?: CandidateId;
  fencingToken?: bigint;
}

export function isAsyncWorkStale(
  identity: AsyncWorkIdentity,
  current: { sessionEpoch?: number; fencingToken?: bigint; candidateId?: string }
): boolean {
  if (current.sessionEpoch !== undefined && identity.sessionEpoch !== current.sessionEpoch) {
    return true;
  }
  if (identity.fencingToken !== undefined && current.fencingToken !== undefined) {
    return identity.fencingToken !== current.fencingToken;
  }
  if (identity.candidateId !== undefined && current.candidateId !== undefined) {
    return identity.candidateId !== current.candidateId;
  }
  return false;
}

/** Errors surfaced by the document runtime. */
export type SessionErrorCode =
  | "lease-held"
  | "candidate-conflict"
  | "source-mutated"
  | "fenced"
  | "session-closed"
  | "candidate-not-ready"
  | "verification-stale"
  | "recovery-required"
  | "unsupported-format"
  | "artifact-missing"
  | "io-error"
  | "policy-denied"
  | "scope-violation";

export class OfficeRuntimeError extends Error {
  constructor(
    readonly code: SessionErrorCode,
    message: string,
    readonly details?: Record<string, unknown>
  ) {
    super(message);
    this.name = "OfficeRuntimeError";
  }
}
