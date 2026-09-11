/**
 * CandidateManager (§37, §66–§72): candidate lifecycle with guarded state
 * transitions, verification hash binding (INV-08) and invalidation on human
 * amendment (INV-09). Agents never mutate committed artifacts (INV-01) —
 * all mutation happens on staging clones owned here.
 */

import type { CandidateId, SessionId } from "../../contracts/ids.js";
import { canTransitionCandidate, type CandidateRevision, type CandidateState } from "../../contracts/candidate.js";
import type { CommittedRevision } from "../../contracts/revision.js";
import type { VerificationReport } from "../../contracts/verification.js";
import { OfficeRuntimeError } from "../../contracts/document.js";
import { newCandidateId } from "../../support/ids.js";
import type { ArtifactStore } from "../../artifact/store/artifact-store.js";
import type { RuntimeRepositories } from "../persistence/repositories.js";
import type { DurableEventBus } from "../sessions/event-bus.js";

export class CandidateManager {
  private readonly candidates = new Map<CandidateId, CandidateRevision>();

  constructor(
    private readonly store: ArtifactStore,
    private readonly repos: RuntimeRepositories,
    private readonly events: DurableEventBus
  ) {}

  get(candidateId: CandidateId): CandidateRevision | undefined {
    const memo = this.candidates.get(candidateId);
    if (memo) return memo;
    const persisted = this.repos.getCandidate(candidateId);
    if (persisted) this.candidates.set(candidateId, persisted);
    return persisted;
  }

  async create(
    sessionId: SessionId,
    sessionEpoch: number,
    base: CommittedRevision,
    createdBy: "agent" | "human"
  ): Promise<CandidateRevision> {
    // Human-amended proposals reuse the same candidate; a brand-new candidate
    // always clones from the committed revision (§66).
    const stagingRef = await this.store.createStagingCopy(base.artifactRef);
    const candidate: CandidateRevision = {
      candidateId: newCandidateId(),
      sessionId,
      baseRevisionId: base.revisionId,
      baseHash: base.contentHash,
      artifactRef: stagingRef,
      state: "preparing",
      createdBy,
      createdAt: Date.now(),
      updatedAt: Date.now()
    };
    this.candidates.set(candidate.candidateId, candidate);
    this.repos.upsertCandidate(candidate);
    await this.events.emit(sessionId, sessionEpoch, "candidate.created", {
      candidateId: candidate.candidateId,
      baseRevisionId: base.revisionId
    });
    return candidate;
  }

  async transition(
    candidateId: CandidateId,
    sessionEpoch: number,
    to: CandidateState,
    patch: Partial<CandidateRevision> = {}
  ): Promise<CandidateRevision> {
    const candidate = this.require(candidateId);
    if (!canTransitionCandidate(candidate.state, to)) {
      throw new OfficeRuntimeError(
        "candidate-not-ready",
        `candidate ${candidateId} cannot transition ${candidate.state} → ${to}`
      );
    }
    const updated: CandidateRevision = {
      ...candidate,
      ...patch,
      state: to,
      updatedAt: Date.now()
    };
    this.candidates.set(candidateId, updated);
    this.repos.upsertCandidate(updated);
    await this.events.emit(candidate.sessionId, sessionEpoch, "candidate.state-changed", {
      candidateId,
      from: candidate.state,
      to
    });
    return updated;
  }

  /**
   * Bind a verification report (INV-08). The report's contentHash must equal
   * the candidate's flushed hash; a candidate without a flush cannot bind a
   * report at all, and a mismatched report is stale (INV-09 path).
   */
  async publishVerification(
    candidateId: CandidateId,
    sessionEpoch: number,
    report: VerificationReport
  ): Promise<CandidateRevision> {
    const candidate = this.require(candidateId);
    if (report.candidateId !== candidateId) {
      throw new OfficeRuntimeError("verification-stale", "report bound to a different candidate");
    }
    if (!candidate.currentHash) {
      throw new OfficeRuntimeError(
        "verification-stale",
        "candidate has no flushed hash; verification cannot bind"
      );
    }
    if (report.contentHash !== candidate.currentHash) {
      throw new OfficeRuntimeError(
        "verification-stale",
        `verification hash ${report.contentHash} does not match candidate hash ${candidate.currentHash}`
      );
    }
    const updated: CandidateRevision = {
      ...candidate,
      verification: report,
      currentHash: report.contentHash,
      updatedAt: Date.now()
    };
    this.candidates.set(candidateId, updated);
    this.repos.upsertCandidate(updated);
    await this.events.emit(candidate.sessionId, sessionEpoch, "candidate.verified", {
      candidateId,
      contentHash: report.contentHash,
      confidence: report.confidence
    });
    return updated;
  }

  /** Human amended the proposal bytes → old verification is void (INV-09). */
  async markHumanAmended(
    candidateId: CandidateId,
    sessionEpoch: number,
    newHash: string
  ): Promise<CandidateRevision> {
    const candidate = this.require(candidateId);
    const updated: CandidateRevision = {
      ...candidate,
      state: candidate.state === "ready" || candidate.state === "verifying" ? "human-amended" : candidate.state,
      currentHash: newHash,
      verification: undefined,
      updatedAt: Date.now()
    };
    this.candidates.set(candidateId, updated);
    this.repos.upsertCandidate(updated);
    if (candidate.verification) {
      await this.events.emit(candidate.sessionId, sessionEpoch, "candidate.verification-invalidated", {
        candidateId,
        previousHash: candidate.currentHash
      });
    }
    return updated;
  }

  async markFailed(
    candidateId: CandidateId,
    sessionEpoch: number,
    reason: string
  ): Promise<CandidateRevision> {
    return this.transition(candidateId, sessionEpoch, "failed", { failureReason: reason });
  }

  async reject(candidateId: CandidateId, sessionEpoch: number): Promise<void> {
    const candidate = this.require(candidateId);
    // Persist the terminal state so DB fallbacks stop reporting it active.
    this.repos.upsertCandidate({
      ...candidate,
      state: "failed",
      failureReason: "rejected",
      updatedAt: Date.now()
    });
    await this.events.emit(candidate.sessionId, sessionEpoch, "candidate.rejected", {
      candidateId
    });
    this.candidates.delete(candidateId);
    await this.store.release(candidate.artifactRef);
  }

  activeForSession(sessionId: SessionId): CandidateRevision | undefined {
    for (const candidate of this.candidates.values()) {
      if (candidate.sessionId === sessionId && candidate.state !== "failed") return candidate;
    }
    return this.repos.activeCandidateForSession(sessionId);
  }

  /** Drop in-memory tracking after finalize; staging bytes released by caller policy. */
  forget(candidateId: CandidateId): void {
    this.candidates.delete(candidateId);
  }

  require(candidateId: CandidateId): CandidateRevision {
    const candidate = this.get(candidateId);
    if (!candidate) {
      throw new OfficeRuntimeError("candidate-not-ready", `unknown candidate: ${candidateId}`);
    }
    return candidate;
  }
}
