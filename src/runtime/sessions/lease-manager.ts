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
import { resolve } from "node:path";
import { canonicalSourceKey } from "../../support/fsx.js";
import type { RuntimeRepositories } from "../persistence/repositories.js";

export class LeaseManager {
  private readonly active = new Map<SessionId, WriterLease>();
  /**
   * P0-4 §38/P11: Single Writer is per LOGICAL DOCUMENT, not per session.
   * Keyed by canonical source path so two sessions on the same file cannot
   * both hold writers (the freeze: "同一个 DocumentSession 最多一个 Writer"
   * with sessions deduplicated per document at the lease level).
   */
  private readonly bySource = new Map<string, WriterLease>();

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
    input: LeaseAcquireInput & { sourcePath: string }
  ): WriterLease {
    const existing = this.active.get(sessionId);
    if (existing) {
      throw new OfficeRuntimeError(
        "lease-held",
        `session ${sessionId} already has an active ${existing.owner} writer (lease ${existing.leaseId})`
      );
    }
    const sourceKey = canonicalSourceKey(input.sourcePath);
    const documentWriter = this.bySource.get(sourceKey);
    if (documentWriter) {
      throw new OfficeRuntimeError(
        "lease-held",
        `document already has an active ${documentWriter.owner} writer from session ${documentWriter.sessionId} (lease ${documentWriter.leaseId})`
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
    // P1 (#6) persistence-first: the DB row is the single commit point. If
    // the insert throws, NO in-memory writer exists — a failed acquire can
    // never leave a phantom lease the runtime treats as active.
    this.repos.insertLease(lease);
    this.active.set(sessionId, lease);
    this.bySource.set(sourceKey, lease);
    return lease;
  }

  release(leaseId: LeaseId): void {
    for (const [sessionId, lease] of this.active) {
      if (lease.leaseId === leaseId) {
        // P1 (#6) persistence-first release: flip the durable audit row
        // BEFORE freeing the in-memory slot — the next writer is only
        // admitted after the DB agrees the previous one ended.
        this.repos.releaseLease(leaseId);
        this.active.delete(sessionId);
        this.removeFromSourceIndex(lease);
        return;
      }
    }
  }

  releaseForSession(sessionId: SessionId): WriterLease | undefined {
    const lease = this.active.get(sessionId);
    if (lease) {
      this.repos.releaseLease(lease.leaseId);
      this.active.delete(sessionId);
      this.removeFromSourceIndex(lease);
    }
    return lease;
  }

  private removeFromSourceIndex(lease: WriterLease): void {
    for (const [key, indexed] of this.bySource) {
      if (indexed.leaseId === lease.leaseId) {
        this.bySource.delete(key);
        return;
      }
    }
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
