/**
 * AgentRuntime (§56, §59–§62, §158): orchestrates the agent write path.
 * Agents never touch committed artifacts (INV-01) and never see real paths
 * (INV-12) — they operate on task contexts bound to candidate staging files.
 */

import type { CandidateId, SessionId } from "../../contracts/ids.js";
import type { MutationCommand, MutationReceipt, MutationScope, OfficeTaskContext } from "../../contracts/capabilities.js";
import type { OfficeEditItem } from "../../contracts/mcp.js";
import { OfficeRuntimeError } from "../../contracts/document.js";
import type { WriterLease } from "../../contracts/lease.js";
import { newCommandId, newReceiptId, newTaskId } from "../../support/ids.js";
import type { ArtifactStore } from "../../artifact/store/artifact-store.js";
import type { ArtifactScanner } from "../../artifact/scanner/scanner.js";
import type { RuntimeRepositories } from "../../runtime/persistence/repositories.js";
import type { SessionManager } from "../../runtime/sessions/session-manager.js";
import type { LeaseManager } from "../../runtime/sessions/lease-manager.js";
import type { CandidateManager } from "../../runtime/candidates/candidate-manager.js";
import type { DurableEventBus } from "../../runtime/sessions/event-bus.js";
import type { Scheduler } from "../../runtime/scheduler/scheduler.js";
import { OperationPolicyEngine } from "./operation-policy.js";
import { OfficeCliAdapter, OfficeCliError } from "./officecli-adapter.js";
import { ResidentPool } from "./resident-pool.js";

export interface AgentRuntimeDeps {
  store: ArtifactStore;
  scanner: ArtifactScanner;
  repos: RuntimeRepositories;
  sessions: SessionManager;
  leases: LeaseManager;
  candidates: CandidateManager;
  events: DurableEventBus;
  scheduler: Scheduler;
  adapter: OfficeCliAdapter;
  pool: ResidentPool;
  policy?: OperationPolicyEngine;
}

export class AgentRuntime {
  readonly policy: OperationPolicyEngine;
  private readonly tasks = new Map<string, { context: OfficeTaskContext; lease: WriterLease; mode: "standalone" | "resident"; residentOpened: boolean }>();

  constructor(private readonly deps: AgentRuntimeDeps) {
    this.policy = deps.policy ?? new OperationPolicyEngine();
  }

  /** Acquire the agent writer lease and clone the committed revision (§158). */
  async beginAgentTask(sessionId: SessionId, scope: MutationScope): Promise<OfficeTaskContext> {
    const session = this.deps.sessions.require(sessionId);
    if (session.lifecycle !== "ready") {
      throw new OfficeRuntimeError("recovery-required", `session lifecycle is ${session.lifecycle}`);
    }
    const epoch = session.sessionEpoch;
    return this.deps.sessions.actor(sessionId).enqueue(async () => {
      const live = this.deps.sessions.require(sessionId);
      if (live.candidate && live.candidate.state !== "failed" && live.candidate.state !== "committing") {
        throw new OfficeRuntimeError(
          "candidate-conflict",
          `session already has an active candidate ${live.candidate.candidateId} in state ${live.candidate.state}`
        );
      }
      const lease = this.deps.leases.acquire(sessionId, epoch, {
        sessionId,
        owner: "agent",
        backend: "officecli",
        baseRevisionId: live.committedRevision.revisionId
      });
      await this.deps.events.emit(sessionId, epoch, "lease.acquired", {
        leaseId: lease.leaseId,
        owner: lease.owner,
        fencingToken: lease.fencingToken.toString()
      });
      const candidate = await this.deps.candidates.create(sessionId, epoch, live.committedRevision, "agent");
      await this.deps.candidates.transition(candidate.candidateId, epoch, "mutating");
      this.deps.sessions.updateSession(sessionId, (s) => {
        s.writerLease = lease;
        s.candidate = candidate;
      });
      const context: OfficeTaskContext = {
        taskId: newTaskId(),
        sessionId,
        documentId: live.documentId,
        candidateId: candidate.candidateId,
        baseRevisionId: live.committedRevision.revisionId,
        fencingToken: lease.fencingToken,
        mutationScope: scope
      };
      this.tasks.set(context.taskId, {
        context,
        lease,
        mode: "resident",
        residentOpened: false
      });
      return context;
    });
  }

  /**
   * Execute an idempotent mutating command against the candidate (§61).
   * Same idempotencyKey replays the stored receipt instead of re-executing.
   */
  async executeMutation(
    task: OfficeTaskContext,
    command: Omit<MutationCommand<OfficeEditItem[]>, "candidateId" | "fencingToken"> & { approved?: boolean }
  ): Promise<MutationReceipt> {
    const tracked = this.tasks.get(task.taskId);
    if (!tracked) {
      throw new OfficeRuntimeError("recovery-required", `unknown or finalized task ${task.taskId}`);
    }
    // Idempotent replay (INV-04).
    const stored = this.deps.repos.getIdempotentReceipt(task.sessionId, command.idempotencyKey);
    if (stored) {
      if (stored.commandId !== command.commandId) {
        return stored.receipt;
      }
      return stored.receipt;
    }

    // Fencing check (INV-03).
    this.deps.leases.validate(tracked.lease, task.fencingToken);

    // Policy + scope evaluation for every item (§124, §60).
    const approved = command.approved === true;
    const decisions = command.payload.map((item) =>
      this.policy.evaluate(item, task.mutationScope, approved)
    );
    const denied = decisions.find((d) => d.risk === "denied");
    if (denied) {
      throw new OfficeRuntimeError(
        "policy-denied",
        denied.reason ?? "policy denied",
        { command: denied.item.command, path: denied.item.path }
      );
    }
    const unapproved = decisions.find((d) => d.requiresApproval);
    if (unapproved) {
      throw new OfficeRuntimeError(
        "policy-denied",
        `operation requires explicit approval: ${unapproved.reason ?? unapproved.risk}`,
        { risk: unapproved.risk }
      );
    }

    const candidate = this.deps.candidates.get(task.candidateId);
    if (!candidate) throw new OfficeRuntimeError("candidate-not-ready", "candidate missing");
    const candidatePath = this.deps.store.resolvePath(candidate.artifactRef);

    const lease = await this.deps.scheduler
      .submit({
        label: `officecli-batch:${task.candidateId}`,
        priority: "AGENT_FOREGROUND",
        identity: {
          sessionId: task.sessionId,
          sessionEpoch: this.deps.sessions.require(task.sessionId).sessionEpoch,
          artifactRef: candidate.artifactRef,
          candidateId: task.candidateId,
          fencingToken: task.fencingToken
        },
        current: () => {
          const session = this.deps.sessions.get(task.sessionId);
          if (!session) return undefined;
          return {
            sessionEpoch: session.sessionEpoch,
            candidateId: session.candidate?.candidateId,
            fencingToken: session.writerLease?.fencingToken
          };
        },
        run: async () => {
          if (tracked.mode === "resident") {
            if (!tracked.residentOpened) {
              await this.deps.pool.acquire(candidatePath);
              tracked.residentOpened = true;
            }
            return this.deps.adapter.runBatchStandalone(candidatePath, command.payload);
          }
          return this.deps.adapter.runBatchStandalone(candidatePath, command.payload);
        }
      })
      .promise;

    const affected = command.payload
      .map((item) => item.path ?? item.parent ?? item.selector ?? item.part ?? item.type ?? item.command)
      .filter((t): t is string => typeof t === "string");
    const receipt: MutationReceipt = {
      receiptId: newReceiptId(),
      commandId: command.commandId,
      candidateId: task.candidateId,
      engine: "officecli",
      affectedTargets: affected,
      completedAt: Date.now()
    };
    this.deps.repos.saveIdempotentReceipt(
      task.sessionId,
      command.idempotencyKey,
      command.commandId,
      receipt
    );
    void lease;
    return receipt;
  }

  /**
   * Flush barrier (§64, INV-06/07): save makes mutations disk-visible, close
   * finalizes the engine's bytes (officecli's close-flush may differ from
   * save-flush), then the strong scan snapshots the authoritative hash the
   * verification binds to. The resident is released here (§65) since the
   * candidate moves to verification/review.
   */
  async flushCandidate(task: OfficeTaskContext): Promise<string> {
    const tracked = this.requireTask(task);
    this.deps.leases.validate(tracked.lease, task.fencingToken);
    const session = this.deps.sessions.require(task.sessionId);
    const candidate = this.deps.candidates.require(task.candidateId);
    const candidatePath = this.deps.store.resolvePath(candidate.artifactRef);

    await this.deps.candidates.transition(task.candidateId, session.sessionEpoch, "flushing");
    if (tracked.residentOpened) {
      await this.deps.adapter.save(candidatePath); // read visibility barrier
      await this.deps.pool.evict(candidatePath); // writer handoff: final bytes
      tracked.residentOpened = false;
    }
    const contentHash = await this.deps.scheduler
      .submit({
        label: `candidate-scan:${task.candidateId}`,
        priority: "VERIFICATION",
        run: async () => {
          const scan = await this.deps.scanner.scan(candidatePath);
          if (!scan.integrity.ok) {
            throw new OfficeRuntimeError(
              "io-error",
              `candidate package integrity failed: ${scan.integrity.corruptEntries.join(", ")}`
            );
          }
          return scan.contentHash;
        }
      })
      .promise;

    await this.deps.candidates.transition(task.candidateId, session.sessionEpoch, "verifying", {
      currentHash: contentHash
    });
    return contentHash;
  }

  /**
   * Writer handoff barrier (§64, INV-07): close the resident and release the
   * agent lease so humans can review (and optionally amend) the proposal.
   */
  async finalizeAgentTask(task: OfficeTaskContext): Promise<void> {
    const tracked = this.requireTask(task);
    this.deps.leases.validate(tracked.lease, task.fencingToken);
    const session = this.deps.sessions.require(task.sessionId);
    const candidate = this.deps.candidates.require(task.candidateId);
    const candidatePath = this.deps.store.resolvePath(candidate.artifactRef);

    if (tracked.residentOpened) {
      await this.deps.pool.evict(candidatePath);
      tracked.residentOpened = false;
    }
    this.deps.leases.release(tracked.lease.leaseId);
    await this.deps.events.emit(task.sessionId, session.sessionEpoch, "lease.released", {
      leaseId: tracked.lease.leaseId,
      owner: "agent"
    });
    this.deps.sessions.updateSession(task.sessionId, (s) => {
      s.writerLease = undefined;
    });
    this.tasks.delete(task.taskId);
  }

  /** Resident on the candidate — exposed for memory telemetry/tests (§65). */
  residentCount(): number {
    return this.deps.pool.size();
  }

  /** Find the live task bound to a candidate (MCP edit reuses it across calls). */
  taskForCandidate(candidateId: CandidateId): OfficeTaskContext | undefined {
    for (const tracked of this.tasks.values()) {
      if (tracked.context.candidateId === candidateId) return tracked.context;
    }
    return undefined;
  }

  activeTaskCount(): number {
    return this.tasks.size;
  }

  async dispose(): Promise<void> {
    for (const [taskId, tracked] of [...this.tasks]) {
      try {
        await this.finalizeAgentTask(tracked.context);
      } catch {
        // Best-effort release during teardown.
        this.deps.leases.release(tracked.lease.leaseId);
      }
      this.tasks.delete(taskId);
    }
    await this.deps.pool.dispose();
  }

  private requireTask(task: OfficeTaskContext): {
    context: OfficeTaskContext;
    lease: WriterLease;
    mode: "standalone" | "resident";
    residentOpened: boolean;
  } {
    const tracked = this.tasks.get(task.taskId);
    if (!tracked) {
      throw new OfficeRuntimeError("recovery-required", `unknown or finalized task ${task.taskId}`);
    }
    return tracked;
  }
}
