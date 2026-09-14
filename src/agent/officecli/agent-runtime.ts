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

/** Stable payload digest for idempotency conflict detection (§61). */
function digestPayload(items: unknown): string {
  const json = JSON.stringify(items);
  let hash = 0x811c9dc5;
  for (let i = 0; i < json.length; i++) {
    hash ^= json.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return `${json.length}:${(hash >>> 0).toString(36)}`;
}

export class AgentRuntime {
  readonly policy: OperationPolicyEngine;
  private readonly tasks = new Map<string, { context: OfficeTaskContext; lease: WriterLease; mode: "standalone" | "resident"; residentOpened: boolean }>();
  /**
   * P0 (#6, INV-04): per-candidate mutation serial lane. The FULL critical
   * section — idempotency lookup, payload-digest check, lease/policy gate,
   * engine mutation, durable receipt persist — runs one-at-a-time per
   * candidate. Concurrent same-key retries coalesce onto the first call's
   * result (the joiner re-reads the persisted receipt), and same-key with a
   * DIFFERENT payload conflicts before that caller's engine side effect.
   * `INSERT OR IGNORE` alone cannot provide this — DB uniqueness only lands
   * after the side effect.
   */
  private readonly mutationLanes = new Map<string, Promise<unknown>>();

  constructor(private readonly deps: AgentRuntimeDeps) {
    this.policy = deps.policy ?? new OperationPolicyEngine();
  }

  /**
   * Acquire the agent writer lease and clone the committed revision (§158).
   * P0-6 three-phase: the mailbox is held only for control; the file clone
   * runs on the scheduler and the candidate registers back inside the actor
   * only if the revision is still fresh.
   */
  async beginAgentTask(sessionId: SessionId, scope: MutationScope): Promise<OfficeTaskContext> {
    const session = this.deps.sessions.require(sessionId);
    if (session.lifecycle !== "ready") {
      throw new OfficeRuntimeError("recovery-required", `session lifecycle is ${session.lifecycle}`);
    }
    // §52: the agent writer path requires strong identity (open stays light).
    const baseRevision = await this.deps.sessions.ensureStrongIdentity(sessionId);
    const captured = await this.deps.sessions.actor(sessionId).enqueue(async () => {
      const live = this.deps.sessions.require(sessionId);
      if (live.candidate && live.candidate.state !== "failed" && live.candidate.state !== "committing") {
        throw new OfficeRuntimeError(
          "candidate-conflict",
          `session already has an active candidate ${live.candidate.candidateId} in state ${live.candidate.state}`
        );
      }
      return {
        epoch: live.sessionEpoch,
        revisionId: live.committedRevision.revisionId,
        artifactRef: live.artifactRef,
        documentId: live.documentId
      };
    });

    const lease = await this.deps.sessions.actor(sessionId).enqueue(async () => {
      const live = this.deps.sessions.require(sessionId);
      if (live.committedRevision.revisionId !== captured.revisionId || live.sessionEpoch !== captured.epoch) {
        throw new OfficeRuntimeError("recovery-required", "session revision moved; retry the task");
      }
      const acquired = this.deps.leases.acquire(sessionId, live.sessionEpoch, {
        sessionId,
        owner: "agent",
        backend: "officecli",
        baseRevisionId: live.committedRevision.revisionId,
        sourcePath: this.deps.store.resolvePath(live.artifactRef)
      });
      await this.deps.events.emit(sessionId, live.sessionEpoch, "lease.acquired", {
        leaseId: acquired.leaseId,
        owner: acquired.owner,
        fencingToken: acquired.fencingToken.toString()
      });
      this.deps.sessions.updateSession(sessionId, (s) => {
        s.writerLease = acquired;
      });
      return acquired;
    });

    // Clone dispatched OFF the actor mailbox (heavy I/O, §42).
    let stagingRef: string;
    try {
      stagingRef = await this.deps.scheduler
        .submit({
          label: `candidate-clone:${sessionId}`,
          priority: "AGENT_FOREGROUND",
          resources: { io: 1 },
          run: () => this.deps.store.createStagingCopy(captured.artifactRef)
        })
        .promise;
    } catch (error) {
      // Roll the lease back; the mailbox was free during the clone.
      this.deps.leases.release(lease.leaseId);
      throw error;
    }

    return this.deps.sessions.actor(sessionId).enqueue(async () => {
      const live = this.deps.sessions.require(sessionId);
      // P0-B four-way gate at the final apply: lifecycle must still be ready.
      if (live.lifecycle !== "ready") {
        this.deps.leases.release(lease.leaseId);
        this.deps.sessions.updateSession(sessionId, (s) => {
          s.writerLease = undefined;
        });
        throw new OfficeRuntimeError("recovery-required", `session lifecycle is ${live.lifecycle}; task aborted`);
      }
      if (live.committedRevision.revisionId !== captured.revisionId || live.sessionEpoch !== captured.epoch) {
        this.deps.leases.release(lease.leaseId);
        this.deps.sessions.updateSession(sessionId, (s) => {
          s.writerLease = undefined;
        });
        throw new OfficeRuntimeError("recovery-required", "session revision moved during clone; retry the task");
      }
      const candidate = await this.deps.candidates.register(
        sessionId,
        live.sessionEpoch,
        { ...baseRevision, artifactRef: captured.artifactRef },
        stagingRef,
        "agent"
      );
      await this.deps.candidates.transition(candidate.candidateId, live.sessionEpoch, "mutating");
      this.deps.sessions.updateSession(sessionId, (s) => {
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
   * P0 (#6): the whole section runs on the candidate's serial lane — see
   * `mutationLanes`.
   */
  async executeMutation(
    task: OfficeTaskContext,
    command: Omit<MutationCommand<OfficeEditItem[]>, "candidateId" | "fencingToken"> & { approved?: boolean }
  ): Promise<MutationReceipt> {
    const tracked = this.tasks.get(task.taskId);
    if (!tracked) {
      throw new OfficeRuntimeError("recovery-required", `unknown or finalized task ${task.taskId}`);
    }
    const payloadDigest = digestPayload(command.payload);

    // Serialize per candidate in CALL order — deterministic command ordering
    // for different keys, and same-key retries see the just-persisted
    // receipt of the first call (one engine execution, identical receipts).
    const previous = this.mutationLanes.get(task.candidateId) ?? Promise.resolve();
    const execution = previous
      .catch(() => undefined)
      .then(() => this.executeMutationSection(task, command, payloadDigest, tracked));
    this.mutationLanes.set(
      task.candidateId,
      execution.then(
        () => undefined,
        () => undefined
      )
    );
    return execution;
  }

  /** The idempotency → policy → engine → persist critical section (P0 #6). */
  private async executeMutationSection(
    task: OfficeTaskContext,
    command: Omit<MutationCommand<OfficeEditItem[]>, "candidateId" | "fencingToken"> & { approved?: boolean },
    payloadDigest: string,
    tracked: { context: OfficeTaskContext; lease: WriterLease; mode: "standalone" | "resident"; residentOpened: boolean }
  ): Promise<MutationReceipt> {
    // Idempotent replay (INV-04) — CANDIDATE-scoped (§61): retries bind to
    // the task's candidate, never to the session lifetime. Same key with a
    // different payload digest is a conflict, not a silent replay — and on
    // the serial lane the conflict is detected before THIS caller mutates.
    const stored = this.deps.repos.getIdempotentReceipt(task.candidateId, command.idempotencyKey);
    if (stored) {
      if (stored.payloadDigest !== payloadDigest) {
        throw new OfficeRuntimeError(
          "idempotency-conflict",
          `idempotencyKey '${command.idempotencyKey}' was already used with a different payload on candidate ${task.candidateId}`
        );
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
            // Explicit resident semantics (§64): in-memory apply through the
            // live resident; flush deferred to save/close.
            return this.deps.adapter.runBatchResident(candidatePath, command.payload);
          }
          // Standalone policy (§63): one-shot open/execute/save cycle.
          if (tracked.residentOpened) {
            await this.deps.pool.evict(candidatePath);
            tracked.residentOpened = false;
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
    this.deps.repos.saveIdempotentReceipt({
      candidateId: task.candidateId,
      idempotencyKey: command.idempotencyKey,
      commandId: command.commandId,
      payloadDigest,
      receipt
    });
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

  /**
   * P1-high (#6): session teardown for the agent write plane. closeSession
   * calls this BEFORE the generic human-edit cleanup so an active task is
   * fully reclaimed — resident evicted, agent lease released with the TRUE
   * owner in the durable event, task/lane entries removed, and the abandoned
   * candidate transitioned to an explicit observable `failed` state instead
   * of lingering ownerless.
   */
  async abortSession(sessionId: SessionId, reason = "session-closed"): Promise<void> {
    for (const [taskId, tracked] of [...this.tasks]) {
      if (tracked.context.sessionId !== sessionId) continue;
      const candidate = this.deps.candidates.get(tracked.context.candidateId);
      const candidatePath = candidate ? this.deps.store.resolvePath(candidate.artifactRef) : undefined;
      if (tracked.residentOpened && candidatePath) {
        await this.deps.pool.evict(candidatePath).catch(() => undefined);
        tracked.residentOpened = false;
      }
      const session = this.deps.sessions.get(sessionId);
      const epoch = session?.sessionEpoch ?? tracked.lease.sessionEpoch;
      if (candidate && candidate.state !== "failed" && candidate.state !== "committing") {
        await this.deps.candidates
          .markFailed(tracked.context.candidateId, epoch, `session closed: ${reason}`)
          .catch(() => undefined);
      }
      this.deps.leases.release(tracked.lease.leaseId);
      await this.deps.events
        .emit(sessionId, epoch, "lease.released", {
          leaseId: tracked.lease.leaseId,
          owner: "agent",
          reason
        })
        .catch(() => undefined);
      if (session) {
        this.deps.sessions.updateSession(sessionId, (s) => {
          s.writerLease = undefined;
          if (s.candidate?.candidateId === tracked.context.candidateId) s.candidate = undefined;
        });
      }
      this.mutationLanes.delete(tracked.context.candidateId);
      this.tasks.delete(taskId);
    }
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

  /** Live task contexts (RPC dispatch + host introspection). */
  activeTasks(): OfficeTaskContext[] {
    return [...this.tasks.values()].map((tracked) => tracked.context);
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
