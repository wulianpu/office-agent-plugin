/**
 * WriterLease management (§38–§40, P11): at most one writer per session;
 * fencing tokens are monotonic per session and persisted so they survive
 * restarts (INV-02, INV-03).
 */

import type { LeaseId, SessionId } from "../../contracts/ids.js";
import type { LeaseAcquireInput, WriterLease } from "../../contracts/lease.js";
import { FencedError } from "../../contracts/lease.js";
import { OfficeRuntimeError } from "../../contracts/document.js";
import { newLeaseId } from "../../support/ids.js";
import type { RuntimeRepositories } from "../persistence/repositories.js";

export class LeaseManager {
  private readonly active = new Map<SessionId, WriterLease>();

  constructor(private readonly repos: RuntimeRepositories) {}

  activeLease(sessionId: SessionId): WriterLease | undefined {
    return this.active.get(sessionId);
  }

  hasActiveLease(sessionId: SessionId): boolean {
    return this.active.has(sessionId);
  }

  acquire(
    sessionId: SessionId,
    sessionEpoch: number,
    input: LeaseAcquireInput
  ): WriterLease {
    const existing = this.active.get(sessionId);
    if (existing) {
      throw new OfficeRuntimeError(
        "lease-held",
        `session ${sessionId} already has an active ${existing.owner} writer (lease ${existing.leaseId})`
      );
    }
    const fencingToken = this.repos.latestFencingToken(sessionId) + 1n;
    const lease: WriterLease = {
      leaseId: newLeaseId(),
      sessionId,
      owner: input.owner,
      backend: input.backend,
      baseRevisionId: input.baseRevisionId,
      fencingToken,
      sessionEpoch,
      acquiredAt: Date.now()
    };
    this.active.set(sessionId, lease);
    this.repos.insertLease(lease);
    return lease;
  }

  release(leaseId: LeaseId): void {
    for (const [sessionId, lease] of this.active) {
      if (lease.leaseId === leaseId) {
        this.active.delete(sessionId);
        this.repos.releaseLease(leaseId);
        return;
      }
    }
  }

  releaseForSession(sessionId: SessionId): WriterLease | undefined {
    const lease = this.active.get(sessionId);
    if (lease) {
      this.active.delete(sessionId);
      this.repos.releaseLease(lease.leaseId);
    }
    return lease;
  }

  /** Validate a writer action against the current lease (INV-03). */
  validate(lease: WriterLease, fencingToken: bigint): void {
    const current = this.active.get(lease.sessionId);
    if (!current || current.leaseId !== lease.leaseId) {
      throw new FencedError(current?.fencingToken ?? -1n, fencingToken, lease.leaseId);
    }
    if (current.fencingToken !== fencingToken) {
      throw new FencedError(current.fencingToken, fencingToken, lease.leaseId);
    }
  }

  /** Rehydrate recovered leases from persistence (crash recovery). */
  rehydrate(): void {
    // Active in-memory leases cannot survive a process crash; DB rows keep an
    // audit trail. Any lease recorded without release is considered abandoned.
  }
}
