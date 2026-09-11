/**
 * OfficeRuntimeService (§6–§8): composition root. Owns the SQLite DB,
 * ArtifactStore/Registry, sessions, candidates, leases, commit machinery,
 * agent runtime, verification and preview — and exposes the complete
 * Preview → Open → Edit → Agent → Review → Accept lifecycle (§157–§158).
 */

import { mkdir } from "node:fs/promises";
import { join } from "node:path";
import type {
  ArtifactLease,
  ArtifactScanResult
} from "../../contracts/artifact.js";
import type { DocumentSession, ViewBookmark } from "../../contracts/document.js";
import type { CandidateRevision } from "../../contracts/candidate.js";
import type { MutationCommand, MutationScope, OfficeTaskContext } from "../../contracts/capabilities.js";
import type { OfficeEditItem } from "../../contracts/mcp.js";
import type { PreviewRequest, PreviewResult } from "../../contracts/preview.js";
import type { VerificationReport } from "../../contracts/verification.js";
import type { EditorInstance, OfficeEditorPlugin } from "../../contracts/editor.js";
import type { WriterLease } from "../../contracts/lease.js";
import { OfficeRuntimeError } from "../../contracts/document.js";
import { ArtifactStore } from "../../artifact/store/artifact-store.js";
import { ArtifactScanner } from "../../artifact/scanner/scanner.js";
import { ArtifactRegistry } from "../../artifact/registry/artifact-registry.js";
import { BasicFormatRuntime } from "../../artifact/runtime/basic-format-runtime.js";
import {
  GenOfficeDocxFormatRuntime,
  GenOfficePptxFormatRuntime
} from "../../artifact/runtime/genoffice-format-runtime.js";
import { probeVendorEngines } from "../../vendor/genoffice/wrapper.js";
import { XlsxSidecarClient, sidecarPreviewWindow } from "../../vendor/genoffice/xlsx-sidecar.js";
import { RuntimeDatabase } from "../persistence/database.js";
import { RuntimeRepositories } from "../persistence/repositories.js";
import { Scheduler } from "../scheduler/scheduler.js";
import { ResourceGovernor, DEFAULT_BUDGETS, type ResourceBudgets } from "../resources/resource-governor.js";
import { LeaseManager } from "../sessions/lease-manager.js";
import { DurableEventBus, EditorSignalBusImpl } from "../sessions/event-bus.js";
import { SessionManager } from "../sessions/session-manager.js";
import { RevisionLog } from "../revisions/revision-log.js";
import { CandidateManager } from "../candidates/candidate-manager.js";
import { AtomicFileCommitter } from "../commit/atomic-file-committer.js";
import { SelfWriteGuardRegistry, SourceWatcher } from "../commit/self-write-guard.js";
import { RecoveryService } from "../recovery/recovery-service.js";
import { promoteToEdit } from "../sessions/edit-promotion.js";
import { OfficeCliAdapter } from "../../agent/officecli/officecli-adapter.js";
import { ResidentPool } from "../../agent/officecli/resident-pool.js";
import { AgentRuntime } from "../../agent/officecli/agent-runtime.js";
import { OperationPolicyEngine } from "../../agent/officecli/operation-policy.js";
import { VerificationPipeline } from "../../verification/pipeline.js";
import { PreviewService } from "../../preview/preview-service.js";
import { RasterCache } from "../../preview/raster.js";
import { WpsHostAdapter } from "../../hosts/wps/wps-adapter.js";
import { InProcessEditorHost } from "../../editors/common/editor-host.js";
import { createBasicEditorPlugin } from "../../editors/common/basic-plugin.js";
import type { RecoveryOutcome } from "../../contracts/revision.js";
import type { ArtifactRef } from "../../contracts/ids.js";
import { sha256File } from "../../support/fsx.js";

export interface OfficeRuntimeServiceOptions {
  workspaceRoot: string;
  dbPath?: string;
  budgets?: Partial<ResourceBudgets>;
  officecli?: { timeoutMs?: number };
  /** Disable OfficeCLI (offline test mode / engine absent). */
  engineDisabled?: boolean;
  /** Skip the optional WPS COM probe (tests; §86 stays optional). */
  skipHostProbe?: boolean;
}

export class OfficeRuntimeService {
  // Infrastructure
  readonly db: RuntimeDatabase;
  readonly repos: RuntimeRepositories;
  readonly store: ArtifactStore;
  readonly scanner: ArtifactScanner;
  readonly registry: ArtifactRegistry;
  readonly governor: ResourceGovernor;
  readonly scheduler: Scheduler;
  readonly events: DurableEventBus;
  readonly editorSignals = new EditorSignalBusImpl();

  // Document runtime
  readonly revisions: RevisionLog;
  readonly candidates: CandidateManager;
  readonly leases: LeaseManager;
  readonly sessions: SessionManager;

  // Commit machinery
  readonly selfWrites = new SelfWriteGuardRegistry();
  readonly watcher: SourceWatcher;
  readonly committer: AtomicFileCommitter;
  readonly recovery: RecoveryService;

  // Agent runtime
  readonly officecli: OfficeCliAdapter;
  readonly residentPool: ResidentPool;
  readonly agent: AgentRuntime;

  // Read path
  readonly previewService: PreviewService;
  readonly verification: VerificationPipeline;

  // Editors
  readonly editorHost = new InProcessEditorHost();

  private engineAvailable = false;
  private genoffice: { pptx: boolean; docx: boolean } = { pptx: false, docx: false };
  private readonly genofficePptxRuntime: GenOfficePptxFormatRuntime;
  private readonly genofficeDocxRuntime: GenOfficeDocxFormatRuntime;
  /** §30: Rust XLSX sidecar (calamine + IronCalc) — bounded-viewport reads. */
  readonly xlsxSidecar = new XlsxSidecarClient();
  /** §91: SVG→PNG rasterizer (L5 pixel renders, host thumbnails). */
  readonly raster = new RasterCache();
  /** §86: optional WPS host adapter (KWPP COM) for L7 certification. */
  readonly wpsHost = new WpsHostAdapter();
  private wpsAvailable = false;
  wpsProbePromise?: Promise<boolean>;
  private disposed = false;
  private readonly editorInstances = new Map<string, EditorInstance>();
  /** ArtifactLeases pinned by live editors — released in endEdit/dispose (P0-2). */
  private readonly editorArtifactLeases = new Map<string, ArtifactLease>();
  /** §11/P11 global: source replacements serialize per physical path, even across sessions. */
  private readonly commitChains = new Map<string, Promise<unknown>>();

  /** Serialize a commit against one source path (multi-session same-file races). */
  private serializeCommit<T>(sourcePath: string, operation: () => Promise<T>): Promise<T> {
    const previous = this.commitChains.get(sourcePath) ?? Promise.resolve();
    const next = previous.then(operation, operation);
    const tail = next
      .catch(() => undefined)
      .finally(() => {
        if (this.commitChains.get(sourcePath) === tail) this.commitChains.delete(sourcePath);
      });
    this.commitChains.set(sourcePath, tail);
    return next;
  }

  private constructor(readonly options: OfficeRuntimeServiceOptions) {
    const budgets: ResourceBudgets = { ...DEFAULT_BUDGETS, ...(options.budgets ?? {}) };
    this.governor = new ResourceGovernor(budgets);
    this.scheduler = new Scheduler({ governor: this.governor });
    this.db = new RuntimeDatabase(options.dbPath ?? join(options.workspaceRoot, "office-runtime.db"));
    this.repos = new RuntimeRepositories(this.db);
    this.events = new DurableEventBus(this.repos);
    this.store = new ArtifactStore(options.workspaceRoot, this.repos);
    this.scanner = new ArtifactScanner(this.repos);
    this.registry = new ArtifactRegistry();
    const resolveByRef = (ref: string) => this.store.resolvePath(ref);
    for (const format of ["docx", "xlsx", "pptx"] as const) {
      const runtime = new BasicFormatRuntime(format);
      runtime.setPathResolver(resolveByRef);
      this.registry.registerRuntime(runtime, "metadata");
    }
    this.registry.setPathResolver(resolveByRef);

    // §146: GenOffice engines replace the basic runtimes for their formats
    // when the vendor bundles are present (probed in initialize()).
    const pptxRuntime = new GenOfficePptxFormatRuntime();
    pptxRuntime.setPathResolver(resolveByRef);
    this.genofficePptxRuntime = pptxRuntime;
    const docxRuntime = new GenOfficeDocxFormatRuntime();
    docxRuntime.setPathResolver(resolveByRef);
    this.genofficeDocxRuntime = docxRuntime;

    this.revisions = new RevisionLog(this.repos);
    this.candidates = new CandidateManager(this.store, this.repos, this.events);
    this.leases = new LeaseManager(this.repos);
    this.sessions = new SessionManager(
      this.store,
      this.scanner,
      this.repos,
      this.events,
      this.revisions,
      this.candidates,
      this.scheduler
    );

    this.watcher = new SourceWatcher(this.selfWrites);
    this.committer = new AtomicFileCommitter(this.repos, this.revisions, this.events, this.selfWrites);
    this.recovery = new RecoveryService(this.repos, this.revisions, this.events, this.committer, this.store);

    this.officecli = new OfficeCliAdapter(options.officecli ?? {});
    this.residentPool = new ResidentPool(this.officecli, this.governor);
    this.agent = new AgentRuntime({
      store: this.store,
      scanner: this.scanner,
      repos: this.repos,
      sessions: this.sessions,
      leases: this.leases,
      candidates: this.candidates,
      events: this.events,
      scheduler: this.scheduler,
      adapter: this.officecli,
      pool: this.residentPool,
      policy: new OperationPolicyEngine()
    });

    this.verification = new VerificationPipeline({
      store: this.store,
      scanner: this.scanner,
      adapter: this.officecli,
      scheduler: this.scheduler,
      raster: this.raster,
      hostAdapter: this.wpsHost,
      hostAdapterAvailable: () => this.wpsAvailable
    });
    this.previewService = new PreviewService(this.store, this.registry, this.scheduler, {
      previewWindow: (path, options) => sidecarPreviewWindow(this.xlsxSidecar, path, options)
    });

    this.registerTrimTargets();
  }

  static async open(options: OfficeRuntimeServiceOptions): Promise<OfficeRuntimeService> {
    await mkdir(options.workspaceRoot, { recursive: true });
    const service = new OfficeRuntimeService(options);
    await service.initialize();
    return service;
  }

  private async initialize(): Promise<void> {
    // P0-1 startup order: DB → store hydrate → commit recovery → session
    // rehydrate. Without hydration every persisted artifactRef resolves to
    // nothing after a process restart.
    await this.store.hydrate();

    if (!this.options.engineDisabled) {
      this.engineAvailable = await this.officecli
        .version_()
        .then(() => true)
        .catch(() => false);
    }
    // §146: register GenOffice runtimes when the vendor bundles load.
    const vendors = await probeVendorEngines();
    this.genoffice = vendors;
    if (vendors.pptx) this.registry.registerRuntime(this.genofficePptxRuntime, "full");
    if (vendors.docx) this.registry.registerRuntime(this.genofficeDocxRuntime, "full");

    // §86: probe the optional WPS host in the BACKGROUND — §86 is optional and
    // a COM probe (seconds, retryable under contention) must never gate
    // startup. Tests skip it entirely via skipHostProbe.
    if (!this.options.skipHostProbe) {
      this.wpsProbePromise = WpsHostAdapter.sharedProbe()
        .then((up) => {
          this.wpsAvailable = up;
          return up;
        })
        .catch(() => false);
    }
    // Staging deletion fallback: engine daemons may hold Windows locks.
    this.store.setLockReleaser(async (path) => {
      await this.officecli.close(path).catch(() => undefined);
    });
    this.committer.lockReleaser = async (path) => {
      await this.officecli.close(path).catch(() => undefined);
    };
    // Crash recovery runs before any session work (§77).
    await this.recovery.recoverAll();
    const rehydrated = this.sessions.rehydrateFromDb();
    if (rehydrated > 0) {
      void rehydrated;
    }
    this.watcher.onMutation((event) => {
      void this.handleSourceMutation(event.sourcePath, event.kind);
    });
  }

  // ---- Read path (§157) ----

  async registerArtifact(path: string): Promise<ArtifactRef> {
    return this.store.register(path);
  }

  async preview(request: Omit<PreviewRequest, "requestId"> & { requestId?: string }): Promise<PreviewResult> {
    const { newRequestId } = await import("../../support/ids.js");
    return this.previewService.preview({
      requestId: request.requestId ?? newRequestId(),
      artifactRef: request.artifactRef,
      priority: request.priority,
      scope: request.scope,
      visual: request.visual
    });
  }

  async acquireArtifactLease(
    artifactRef: ArtifactRef,
    consumer: string,
    profile: "metadata" | "full" = "metadata"
  ): Promise<ArtifactLease> {
    return this.registry.acquire({
      artifactRef,
      format: this.store.formatOf(artifactRef),
      consistency: "optimistic",
      profile,
      priority: "VISIBLE_PREVIEW",
      consumer
    });
  }

  // ---- Sessions (§12–§13) ----

  async openSession(artifactRef: ArtifactRef): Promise<DocumentSession> {
    // P1-high: install the watcher BEFORE open — the background strong hash
    // starts inside open(), and a mutation in that gap must not be missed.
    const sourcePath = this.store.resolvePath(artifactRef);
    this.watcher.watchFile(sourcePath);
    try {
      return await this.sessions.open(artifactRef);
    } catch (error) {
      this.watcher.unwatchFile(sourcePath);
      throw error;
    }
  }

  getSession(sessionId: string): DocumentSession | undefined {
    return this.sessions.get(sessionId);
  }

  /**
   * §77 resolution path for sessions rehydrated in `recovery-required`:
   * filesystem hash facts decide — match with the latest committed revision
   * restores `ready`; anything else is an explicit `conflict`.
   */
  async resolveRecoveredSession(sessionId: string): Promise<"ready" | "conflict"> {
    const session = this.sessions.require(sessionId);
    if (session.lifecycle !== "recovery-required") {
      throw new OfficeRuntimeError("recovery-required", `session is ${session.lifecycle}, not recovery-required`);
    }
    // P0-6 three-phase: capture → hash outside the actor → apply if fresh.
    const captured = await this.sessions.actor(sessionId).enqueue(async () => {
      const live = this.sessions.require(sessionId);
      return { epoch: live.sessionEpoch, revisionId: live.committedRevision.revisionId, artifactRef: live.artifactRef };
    });
    const sourcePath = this.store.resolvePath(captured.artifactRef);
    const currentHash = await this.scheduler
      .submit({ label: `recovery-hash:${sessionId}`, priority: "EDIT_PROMOTION", run: () => sha256File(sourcePath) })
      .promise;
    return this.sessions.actor(sessionId).enqueue(async () => {
      const live = this.sessions.require(sessionId);
      if (live.sessionEpoch !== captured.epoch || live.committedRevision.revisionId !== captured.revisionId) {
        throw new OfficeRuntimeError("recovery-required", "session state changed during recovery resolution; retry");
      }
      if (currentHash === live.committedRevision.contentHash) {
        this.sessions.updateSession(sessionId, (s) => {
          s.lifecycle = "ready";
        });
        return "ready";
      }
      this.sessions.updateSession(sessionId, (s) => {
        s.lifecycle = "conflict";
      });
      await this.events.emit(sessionId, live.sessionEpoch, "session.conflict", {
        reason: "recovery-hash-mismatch",
        expected: live.committedRevision.contentHash,
        found: currentHash
      });
      return "conflict";
    });
  }

  /** §13: first mutation promotes the read-only tab to an editor in place. */
  async beginEdit(sessionId: string, bookmark?: ViewBookmark): Promise<{ lease: WriterLease; editor: EditorInstance }> {
    const session = this.sessions.require(sessionId);
    const result = await promoteToEdit(
      {
        store: this.store,
        scanner: this.scanner,
        repos: this.repos,
        sessions: this.sessions,
        leases: this.leases,
        candidates: this.candidates,
        events: this.events,
        scheduler: this.scheduler
      },
      sessionId,
      bookmark
    );
    const plugin: OfficeEditorPlugin | undefined = this.editorHost.pluginFor(session.format);
    if (!plugin) {
      throw new OfficeRuntimeError("unsupported-format", `no editor plugin for ${session.format}`);
    }
    const artifactLease = await this.acquireArtifactLease(session.artifactRef, `editor:${sessionId}`, "full");
    const editor = await this.editorHost.mount(plugin, {
      sessionId,
      artifactContext: artifactLease.context,
      bookmark,
      readOnly: false
    });
    await editor.activateEdit();
    this.sessions.updateSession(sessionId, (s) => {
      s.editor = {
        instanceId: editor.instanceId,
        plugin,
        host: this.editorHost,
        boundAt: Date.now(),
        artifactLease
      };
    });
    this.editorInstances.set(sessionId, editor);
    this.editorArtifactLeases.set(sessionId, artifactLease);
    return { lease: result.lease, editor };
  }

  /**
   * Human save (Ctrl+S): commits a human revision through the atomic
   * committer when the editor flushed new bytes, otherwise rebinds if the
   * source changed directly.
   */
  async humanSave(sessionId: string, newContentArtifactRef?: ArtifactRef): Promise<{ revisionId: string; unchanged: boolean }> {
    await this.sessions.ensureStrongIdentity(sessionId);
    const session = this.sessions.require(sessionId);
    if (!this.leases.activeLease(sessionId) || this.leases.activeLease(sessionId)!.owner !== "human") {
      throw new OfficeRuntimeError("lease-held", "human save requires the human writer lease");
    }
    const sourcePath = this.store.resolvePath(session.artifactRef);
    const currentHash = await this.scanner.hashOnly(sourcePath);

    if (!newContentArtifactRef) {
      // P0-5: a bare save validates the source is still at the committed
      // revision. Editors NEVER write the source in place — content goes to a
      // staging artifact and through the atomic committer (same crash
      // semantics as agents). A mutated source here is a conflict, not a
      // revision (§74: all source replacement flows through the committer).
      if (currentHash !== session.committedRevision.contentHash) {
        this.sessions.updateSession(sessionId, (s) => {
          s.lifecycle = s.lifecycle === "ready" ? "conflict" : s.lifecycle;
        });
        await this.events.emit(sessionId, session.sessionEpoch, "session.conflict", {
          reason: "source-mutated-on-save",
          expected: session.committedRevision.contentHash,
          found: currentHash
        });
        throw new OfficeRuntimeError(
          "source-mutated",
          "source changed outside the editor; save requires a staging artifact or a reopen"
        );
      }
      return { revisionId: session.committedRevision.revisionId, unchanged: true };
    }

    const candidatePath = this.store.resolvePath(newContentArtifactRef);
    const candidateHash = await this.scanner.hashOnly(candidatePath);
    if (candidateHash === currentHash) {
      return { revisionId: session.committedRevision.revisionId, unchanged: true };
    }
    const result = await this.serializeCommit(sourcePath, () =>
      this.committer.commit({
        sessionId,
        sessionEpoch: session.sessionEpoch,
        candidateId: `human-${sessionId}`,
        sourcePath,
        candidatePath,
        sessionArtifactRef: session.artifactRef,
        expectedSourceHash: session.committedRevision.contentHash,
        candidateHash,
        origin: "human"
      })
    );
    this.bindCommittedRevision(sessionId, result.newRevision);
    return { revisionId: result.newRevision.revisionId, unchanged: false };
  }

  /** Release the human lease (editor closed without saving / after save). */
  async endEdit(sessionId: string): Promise<void> {
    const editor = this.editorInstances.get(sessionId);
    if (editor) {
      await editor.dispose().catch(() => undefined);
      this.editorInstances.delete(sessionId);
    }
    this.editorArtifactLeases.get(sessionId)?.release();
    this.editorArtifactLeases.delete(sessionId);
    const lease = this.leases.releaseForSession(sessionId);
    if (lease) {
      await this.events.emit(sessionId, this.sessions.require(sessionId).sessionEpoch, "lease.released", {
        leaseId: lease.leaseId,
        owner: "human"
      });
    }
    this.sessions.updateSession(sessionId, (s) => {
      s.writerLease = undefined;
      s.editor = undefined;
    });
  }

  // ---- Agent write path (§158) ----

  async beginAgentTask(sessionId: string, scope: MutationScope): Promise<OfficeTaskContext> {
    return this.agent.beginAgentTask(sessionId, scope);
  }

  async executeAgentMutation(
    task: OfficeTaskContext,
    command: Omit<MutationCommand<OfficeEditItem[]>, "candidateId" | "fencingToken"> & { approved?: boolean }
  ) {
    return this.agent.executeMutation(task, command);
  }

  async flushAgentCandidate(task: OfficeTaskContext): Promise<string> {
    const hash = await this.agent.flushCandidate(task);
    this.refreshSessionCandidate(task.sessionId, task.candidateId);
    return hash;
  }

  async verifyAgentCandidate(task: OfficeTaskContext, changedTargets?: string[]): Promise<VerificationReport> {
    const session = this.sessions.require(task.sessionId);
    const candidate = this.candidates.require(task.candidateId);
    if (!candidate.currentHash) {
      throw new OfficeRuntimeError("candidate-not-ready", "candidate not flushed; call flush first");
    }
    const baseRevision = this.revisions.get(candidate.baseRevisionId);
    if (!baseRevision) {
      throw new OfficeRuntimeError("recovery-required", `base revision ${candidate.baseRevisionId} missing`);
    }
    const report = await this.verification.verify({
      candidate,
      baseRevision,
      changedTargets: changedTargets ?? [],
      engineAvailable: this.engineAvailable
    });
    const updated = await this.candidates.publishVerification(
      task.candidateId,
      session.sessionEpoch,
      report
    );
    if (report.confidence === "unverified" || report.structural.status === "fail") {
      await this.candidates.markFailed(
        task.candidateId,
        session.sessionEpoch,
        `verification failed: ${report.structural.issues.map((i) => i.message).join("; ")}`
      );
      throw new OfficeRuntimeError("verification-stale", "candidate failed verification", {
        confidence: report.confidence
      });
    }
    await this.candidates.transition(task.candidateId, session.sessionEpoch, "ready");
    this.refreshSessionCandidate(task.sessionId, task.candidateId);
    void updated;
    return report;
  }

  /** §64: close + writer handoff after verification, before human review. */
  async finalizeAgentTask(task: OfficeTaskContext): Promise<void> {
    await this.agent.finalizeAgentTask(task);
  }

  private refreshSessionCandidate(sessionId: string, candidateId: string): void {
    const candidate = this.candidates.get(candidateId);
    if (candidate) {
      this.sessions.updateSession(sessionId, (s) => {
        s.candidate = candidate;
      });
    }
  }

  /** §158 Accept: fence/base/hash/verify checks, then atomic commit + promotion. */  async acceptCandidate(sessionId: string, candidateId: string): Promise<{ revisionId: string; contextPromoted: boolean }> {
    const session = this.sessions.require(sessionId);
    return this.sessions.actor(sessionId).enqueue(async () => {
      const live = this.sessions.require(sessionId);
      const candidate = this.candidates.require(candidateId);

      if (live.writerLease) {
        throw new OfficeRuntimeError("lease-held", "a writer still holds the session lease");
      }
      if (candidate.state !== "ready" && candidate.state !== "human-amended") {
        throw new OfficeRuntimeError("candidate-not-ready", `candidate state is ${candidate.state}`);
      }
      if (candidate.baseRevisionId !== live.committedRevision.revisionId) {
        throw new OfficeRuntimeError("source-mutated", "committed revision moved past the candidate base");
      }
      if (!candidate.verification || !candidate.currentHash) {
        throw new OfficeRuntimeError("verification-stale", "candidate lacks a bound verification report");
      }
      if (candidate.verification.contentHash !== candidate.currentHash) {
        throw new OfficeRuntimeError("verification-stale", "verification hash does not match candidate");
      }

      const sourcePath = this.store.resolvePath(live.artifactRef);
      const candidatePath = this.store.resolvePath(candidate.artifactRef);

      await this.candidates.transition(candidateId, live.sessionEpoch, "committing");
      const result = await this.serializeCommit(sourcePath, () =>
        this.committer.commit({
          sessionId,
          sessionEpoch: live.sessionEpoch,
          candidateId,
          sourcePath,
          candidatePath,
          sessionArtifactRef: live.artifactRef,
          expectedSourceHash: candidate.baseHash,
          // Non-null: the verification check above guarantees a flushed hash.
          candidateHash: candidate.currentHash!,
          origin: candidate.createdBy
        })
      );

      // Context promotion (§71–§72): same bytes → same parse; only bindings move.
      // The candidate's FULL context (engine model) re-registers under the
      // source's new version key — the next editor acquire hits the cache.
      const stagingPath = this.store.resolvePath(candidate.artifactRef);
      const { fileFingerprint, fingerprintKey } = await import("../../support/fsx.js");
      try {
        const stagingFp = await fileFingerprint(stagingPath);
        this.registry.promote(
          { artifactRef: candidate.artifactRef, fingerprintKey: fingerprintKey(stagingFp), profile: "full" },
          { artifactRef: live.artifactRef, fingerprintKey: fingerprintKey(stagingFp), profile: "full" }
        );
      } catch {
        // Staging file may already be released; promotion is best-effort.
      }
      this.bindCommittedRevision(sessionId, result.newRevision);
      this.candidates.forget(candidateId);
      this.repos.deleteCandidate(candidateId);
      await this.store.release(candidate.artifactRef);
      this.sessions.updateSession(sessionId, (s) => {
        s.candidate = undefined;
      });
      return { revisionId: result.newRevision.revisionId, contextPromoted: true };
    });
  }

  async rejectCandidate(sessionId: string, candidateId: string): Promise<void> {
    const session = this.sessions.require(sessionId);
    await this.candidates.reject(candidateId, session.sessionEpoch);
    this.sessions.updateSession(sessionId, (s) => {
      s.candidate = undefined;
    });
    this.previewService.invalidate(session.artifactRef);
  }

  // ---- Lifecycle ----

  async closeSession(sessionId: string): Promise<void> {
    await this.endEdit(sessionId).catch(() => undefined);
    const session = this.sessions.get(sessionId);
    if (session) {
      this.watcher.unwatchFile(this.store.resolvePath(session.artifactRef));
    }
    await this.sessions.close(sessionId);
  }

  async recoverManually(): Promise<RecoveryOutcome[]> {
    return this.recovery.recoverAll();
  }

  /** Strong scan exposed for hosts/tests (§78). */
  async scanArtifact(artifactRef: ArtifactRef): Promise<ArtifactScanResult> {
    return this.scanner.scan(this.store.resolvePath(artifactRef));
  }

  isEngineAvailable(): boolean {
    return this.engineAvailable;
  }

  /** §146: GenOffice engine availability per format (capability matrix input). */
  isGenOfficeAvailable(format: "pptx" | "docx"): boolean {
    return this.genoffice[format];
  }

  /**
   * §125 live snapshot: the officecli schema fingerprint is derived from the
   * engine's own capability-reference surface (help output digest), so a
   * silent engine upgrade changes the lock the runtime reports.
   */
  async probeCompatibilitySnapshot(): Promise<{
    plugin: string;
    genoffice: string;
    officecli: { version: string; schemaFingerprint: string };
    contract: number;
    dbSchema: number;
  }> {
    const lock = (await import("../../plugin/office-plugin.js")).COMPATIBILITY_LOCK;
    if (!this.engineAvailable) return lock;
    try {
      const { createHash } = await import("node:crypto");
      const help = await this.officecli.run(["help", "--json"]).catch(() => undefined);
      const text = help ? JSON.stringify(help).slice(0, 262_144) : "";
      const version = await this.officecli.version_().catch(() => "unknown");
      const fingerprint = text
        ? createHash("sha256").update(text).digest("hex").slice(0, 16)
        : "unavailable";
      return { ...lock, officecli: { version, schemaFingerprint: fingerprint } };
    } catch {
      return lock;
    }
  }

  /** §30: Rust sidecar availability for XLSX (capability matrix input). */
  isXlsxSidecarAvailable(): boolean {
    return this.xlsxSidecar.available;
  }

  /** §86: optional host adapter availability (capability matrix input). */
  isHostAvailable(id: "wps" | "powerpoint"): boolean {
    return id === "wps" ? this.wpsAvailable : false;
  }

  private bindCommittedRevision(sessionId: string, revision: DocumentSession["committedRevision"]): void {
    this.sessions.updateSession(sessionId, (s) => {
      s.committedRevision = revision;
    });
    const artifactRef = revision.artifactRef;
    this.previewService.invalidate(artifactRef);
    // Backpressure (§112): superseded preview work is dropped.
    this.scheduler.cancelWhere((job) => job.label.startsWith(`preview-render:${artifactRef}`));
  }

  private async handleSourceMutation(sourcePath: string, kind: "self-write" | "external"): Promise<void> {
    if (kind === "self-write") return;
    for (const session of this.sessions.list()) {
      if (this.store.resolvePath(session.artifactRef) !== sourcePath) continue;
      this.sessions.updateSession(session.sessionId, (s) => {
        s.lifecycle = s.lifecycle === "ready" ? "conflict" : s.lifecycle;
      });
      await this.events.emit(session.sessionId, session.sessionEpoch, "session.conflict", {
        reason: "external-mutation"
      });
    }
  }

  private registerTrimTargets(): void {
    // Eviction ladder (§107): lower rank trimmed first.
    this.governor.registerTrimTarget({
      rank: 1,
      label: "expired-previews",
      trim: async () => {
        this.previewService.trimAll();
      }
    });
    this.governor.registerTrimTarget({
      rank: 4,
      label: "idle-residents",
      trim: async () => {
        await this.residentPool.trimIdle();
      }
    });
    this.governor.registerTrimTarget({
      rank: 5,
      label: "scan-cache",
      trim: async () => {
        this.scanner.scans.trimAll();
      }
    });
    this.governor.registerTrimTarget({
      rank: 6,
      label: "cold-artifact-cache",
      trim: async (level) => {
        await this.registry.trim(level);
      }
    });
  }

  async dispose(): Promise<void> {
    if (this.disposed) return;
    this.disposed = true;
    await this.agent.dispose().catch(() => undefined);
    for (const editor of this.editorInstances.values()) {
      await editor.dispose().catch(() => undefined);
    }
    this.editorInstances.clear();
    for (const lease of this.editorArtifactLeases.values()) lease.release();
    this.editorArtifactLeases.clear();
    this.watcher.dispose();
    await this.xlsxSidecar.dispose().catch(() => undefined);
    await this.registry.dispose().catch(() => undefined);
    this.db.close();
  }
}
