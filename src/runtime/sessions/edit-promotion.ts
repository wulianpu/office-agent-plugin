/**
 * Edit Promotion (§13, §51): read-only session → first human mutation promotes
 * the same tab in place into an editor. Checks: stable source identity, no
 * competing candidate, write readiness, then human WriterLease acquisition and
 * editor activation. The triggering operation is replayed by the caller.
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

export async function promoteToEdit(
  deps: EditPromotionDeps,
  sessionId: SessionId,
  bookmark?: ViewBookmark
): Promise<EditPromotionResult> {
  const session = deps.sessions.require(sessionId);
  if (session.lifecycle !== "ready") {
    throw new OfficeRuntimeError("recovery-required", `session lifecycle is ${session.lifecycle}`);
  }
  return deps.sessions.actor(sessionId).enqueue(async () => {
    const live = deps.sessions.require(sessionId);

    // 1. Competing candidate blocks promotion (§13).
    const active = deps.candidates.activeForSession(sessionId);
    if (active && active.state !== "failed") {
      throw new OfficeRuntimeError(
        "candidate-conflict",
        `session has candidate ${active.candidateId} in state ${active.state}; resolve it before editing`
      );
    }

    // 2. Source stability: hash must match the committed revision (§13).
    const sourcePath = deps.store.resolvePath(live.artifactRef);
    const currentHash = await deps.scheduler
      .submit({
        label: `promotion-hash:${sessionId}`,
        priority: "EDIT_PROMOTION",
        run: () => sha256File(sourcePath)
      })
      .promise;
    if (currentHash !== live.committedRevision.contentHash) {
      live.lifecycle = "conflict";
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

    // 3. Write readiness achieved by the hash check above (§52).
    deps.sessions.setReadiness(sessionId, { readiness: "ready", strongHash: currentHash, probedAt: Date.now() });

    // 4. Acquire the human lease (§13).
    const lease = deps.leases.acquire(sessionId, live.sessionEpoch, {
      sessionId,
      owner: "human",
      backend: "genoffice",
      baseRevisionId: live.committedRevision.revisionId
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
