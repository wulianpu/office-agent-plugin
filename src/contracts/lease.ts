/**
 * Phase 0 — Contracts: writer lease & fencing (Design Doc §38–§40, P11).
 */

import type { LeaseId, RevisionId, SessionEpoch, SessionId } from "./ids.js";

export type LeaseOwner = "human" | "agent" | "external";
export type LeaseBackend = "genoffice" | "officecli" | "powerpoint" | "wps";

/**
 * §38: at most one writer per DocumentSession; readers are never blocked (P4, §40).
 * backend "genoffice" covers the plugin's native editor implementations.
 */
export interface WriterLease {
  leaseId: LeaseId;
  sessionId: SessionId;
  owner: LeaseOwner;
  backend: LeaseBackend;
  baseRevisionId: RevisionId;
  /** Monotonic across the session; stale tokens are rejected on completion (§39). */
  fencingToken: bigint;
  sessionEpoch: SessionEpoch;
  acquiredAt: number;
}

export interface LeaseAcquireInput {
  sessionId: SessionId;
  owner: LeaseOwner;
  backend: LeaseBackend;
  baseRevisionId: RevisionId;
}

/** Thrown when a writer action carries a stale fencing token (INV-03). */
export class FencedError extends Error {
  constructor(
    readonly expectedToken: bigint,
    readonly receivedToken: bigint,
    readonly leaseId: LeaseId
  ) {
    super(
      `writer fenced: lease ${leaseId} expected token ${expectedToken}, received ${receivedToken}`
    );
    this.name = "FencedError";
  }
}
