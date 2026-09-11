/**
 * Edit Promotion (§13, §51, P0-6): read-only session → first human mutation
 * promotes the same tab in place into an editor.
 *
 * Three-phase actor discipline (§41–§43): the mailbox is held only for
 * control steps; the strong hash runs outside the actor, and its result is
 * applied back inside the actor only if still fresh.
 *
 *  Actor(capture) → Scheduler(hash) → Actor(apply if fresh)
 */

import type { SessionId } from "../../contracts/ids.js";
import { OfficeRuntimeError } from "../../contracts/document.js";
import type { ViewBookmark } from "../../contracts/document.js";
import type { WriterLease } from "../../contracts/lease.js";
import type { RuntimeRepositories } from "../persistence/repositories.js";
import type { SessionManager } from "./session-manager.js";
import type { LeaseManager } from "./lease-manager.js";
import type { CandidateManager } from "../candidates/candidate-manager.js";
import type { DurableEventBus } from "./event-bus.js";
import type { ArtifactScanner } from "../../artifact/scanner/scanner.js";
import type { ArtifactStore } from "../../artifact/store/artifact-store.js";
import type { Scheduler } from "../scheduler/scheduler.js";
import { sha256File } from "../../support/fsx.js";

export interface EditPromotionDeps {
  store: ArtifactStore;
  scanner: ArtifactScanner;
  repos: RuntimeRepositories;
  sessions: SessionManager;
  leases: LeaseManager;
  candidates: CandidateManager;
  events: DurableEventBus;
  scheduler: Scheduler;
}

export interface EditPromotionResult {
  lease: WriterLease;
  bookmark?: ViewBookmark;
}

interface PromotionCapture {
  epoch: number;
  revisionId: string;
}

export async function promoteToEdit(
  deps: EditPromotionDeps,
  sessionId: SessionId,
  bookmark?: ViewBookmark
): Promise<EditPromotionResult> {
  const session = deps.sessions.require(sessionId);
  if (session.lifecycle !== "ready") {
    throw new OfficeRuntimeError("recovery-required", `session lifecycle is ${session.lifecycle}`);
  }

  // Phase 1 — capture (inside actor, control-only).
  const captured: PromotionCapture = await deps.sessions.actor(sessionId).enqueue(async () => {
    const live = deps.sessions.require(sessionId);
    return {
      epoch: live.sessionEpoch,
      revisionId: live.committedRevision.revisionId
    };
  });

  // Phase 2 — dispatch (outside actor): strong identity + source hash.
  // The mailbox stays free; concurrent control operations can interleave.
  const revision = await deps.sessions.ensureStrongIdentity(sessionId);
  const sourcePath = deps.store.resolvePath(session.artifactRef);
  const currentHash = await deps.scheduler
    .submit({
      label: `promotion-hash:${sessionId}`,
      priority: "EDIT_PROMOTION",
      run: () => sha256File(sourcePath)
    })
    .promise;

  // Phase 3 — apply (inside actor): validate freshness, then decide.
  return deps.sessions.actor(sessionId).enqueue(async () => {
    const live = deps.sessions.require(sessionId);

    // Freshness (§43): state moved during the hash — retryable, not applied.
    if (live.sessionEpoch !== captured.epoch || live.committedRevision.revisionId !== revision.revisionId) {
      throw new OfficeRuntimeError(
        "recovery-required",
        "session state changed during promotion; retry the edit"
      );
    }

    // Competing candidate blocks promotion (§13).
    const active = deps.candidates.activeForSession(sessionId);
    if (active && active.state !== "failed") {
      throw new OfficeRuntimeError(
        "candidate-conflict",
        `session has candidate ${active.candidateId} in state ${active.state}; resolve it before editing`
      );
    }

    // Source stability: hash must match the committed revision (§13).
    if (currentHash !== live.committedRevision.contentHash) {
      deps.sessions.updateSession(sessionId, (s) => {
        s.lifecycle = "conflict";
      });
      await deps.events.emit(sessionId, live.sessionEpoch, "session.conflict", {
        reason: "source-mutated",
        expected: live.committedRevision.contentHash,
        found: currentHash
      });
      throw new OfficeRuntimeError(
        "source-mutated",
        "source file changed underneath the session; reopen to resolve"
      );
    }

    // Write readiness achieved by the hash check above (§52).
    deps.sessions.setReadiness(sessionId, {
      readiness: "ready",
      strongHash: currentHash,
      probedAt: Date.now()
    });

    // Acquire the human lease (§13, P0-4 document scope).
    const lease = deps.leases.acquire(sessionId, live.sessionEpoch, {
      sessionId,
      owner: "human",
      backend: "genoffice",
      baseRevisionId: live.committedRevision.revisionId,
      sourcePath
    });
    await deps.events.emit(sessionId, live.sessionEpoch, "lease.acquired", {
      leaseId: lease.leaseId,
      owner: "human",
      fencingToken: lease.fencingToken.toString()
    });
    deps.sessions.updateSession(sessionId, (s) => {
      s.writerLease = lease;
    });
    return { lease, bookmark };
  });
}
