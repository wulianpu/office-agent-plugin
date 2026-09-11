/**
 * SessionManager + SessionActor (§32–§33, §41–§43): every DocumentSession is
 * a logical actor whose state-changing operations run strictly serially.
 * The actor is control-only — heavy work is dispatched to the Scheduler and
 * validated for freshness before results are applied (§42–§43).
 */

import type { ArtifactRef, DocumentId, OfficeFormat, SessionEpoch, SessionId } from "../../contracts/ids.js";
import type { DocumentSession, WriteReadinessState } from "../../contracts/document.js";
import { OfficeRuntimeError } from "../../contracts/document.js";
import { newDocumentId, newSessionId } from "../../support/ids.js";
import type { ArtifactStore } from "../../artifact/store/artifact-store.js";
import type { ArtifactScanner } from "../../artifact/scanner/scanner.js";
import type { RuntimeRepositories } from "../persistence/repositories.js";
import type { DurableEventBus } from "../sessions/event-bus.js";
import type { CommittedRevision } from "../../contracts/revision.js";
import type { RevisionLog } from "../revisions/revision-log.js";
import type { CandidateManager } from "../candidates/candidate-manager.js";
import { Scheduler } from "../scheduler/scheduler.js";

/**
 * Actor mailbox: enqueued operations execute one-at-a-time per session.
 * Dispatched async jobs (hashing, scanning) run through the scheduler and
 * re-enter the actor for state application.
 */
export class SessionActor {
  private chain: Promise<unknown> = Promise.resolve();

  /** Serialize a control operation on this session (§41). */
  enqueue<T>(operation: () => Promise<T>): Promise<T> {
    const next = this.chain.then(operation, operation);
    this.chain = next.catch(() => undefined);
    return next;
  }
}

export class SessionManager {
  private readonly sessions = new Map<SessionId, { session: DocumentSession; actor: SessionActor; readiness: WriteReadinessState }>();
  /** Background strong-identity promises per session (P0-6 §52). */
  private readonly strongIdentity = new Map<SessionId, Promise<CommittedRevision>>();
  /** P0-B: conflict watcher — installed by the service BEFORE open() so the
   *  background hash never races a missed filesystem event. */
  onSourceMutated?: (sourcePath: string, kind: "self-write" | "external") => void;

  constructor(
    private readonly store: ArtifactStore,
    private readonly scanner: ArtifactScanner,
    private readonly repos: RuntimeRepositories,
    private readonly events: DurableEventBus,
    private readonly revisions: RevisionLog,
    private readonly candidates: CandidateManager,
    private readonly scheduler: Scheduler
  ) {}

  /**
   * Formal open (§12, P0-6/§六): creates a read-only DocumentSession and
   * returns it READY immediately (optimistic fingerprint identity). The
   * strong SHA-256 runs in the background; writer paths (edit/agent/accept)
   * block on `ensureStrongIdentity` — interactive open is never gated by a
   * full-file hash (PERF-03/§52).
   */
  async open(artifactRef: ArtifactRef): Promise<DocumentSession> {
    const format = this.store.formatOf(artifactRef);
    const sessionId = newSessionId();
    const documentId = await this.resolveOrCreateDocumentId(artifactRef);

    this.repos.upsertSession({
      sessionId,
      documentId,
      artifactRef,
      format,
      backend: "managed-file",
      lifecycle: "opening",
      epoch: 1,
      createdAt: Date.now()
    });

    // Optimistic placeholder — strong identity binds in the background.
    const placeholder: CommittedRevision = {
      revisionId: `rev_pending_${sessionId}`,
      sessionId,
      sequence: 0,
      artifactRef,
      contentHash: "",
      origin: "external",
      createdAt: Date.now()
    };

    const session: DocumentSession = {
      sessionId,
      documentId,
      format,
      backend: "managed-file",
      lifecycle: "ready",
      sessionEpoch: 1,
      committedRevision: placeholder,
      artifactRef,
      createdAt: Date.now()
    };
    this.sessions.set(sessionId, { session, actor: new SessionActor(), readiness: { readiness: "probing" } });
    this.repos.upsertSession({
      sessionId,
      documentId,
      artifactRef,
      format,
      backend: "managed-file",
      lifecycle: "ready",
      epoch: 1,
      createdAt: session.createdAt
    });

    // Background strong identity (§52). P0-B hardening:
    //  * watcher installed BEFORE hashing starts (no missed-event window);
    //  * stable hash — fingerprint before/after guards against hybrid reads;
    //  * a closed session never persists an orphan revision.
    const strong = (async () => {
      const sourcePath = this.store.resolvePath(artifactRef);
      const stable = await this.scheduler
        .submit({
          label: `open-hash:${sessionId}`,
          priority: "EDIT_PROMOTION",
          resources: { io: 1 },
          run: async () => {
            const { stableHashFile } = await import("../../support/fsx.js");
            for (let attempt = 0; attempt < 3; attempt++) {
              const result = await stableHashFile(sourcePath);
              if (result) return result;
            }
            return null;
          }
        })
        .promise;
      if (!stable) {
        // File keeps moving under us — surface as a conflict, not a bad hash.
        const live = this.sessions.get(sessionId);
        if (live) {
          this.updateSession(sessionId, (s) => {
            s.lifecycle = s.lifecycle === "ready" ? "conflict" : s.lifecycle;
          });
          await this.events
            .emit(sessionId, 1, "session.conflict", { reason: "unstable-source-identity" })
            .catch(() => undefined);
        }
        throw new OfficeRuntimeError("source-mutated", "source identity never stabilized during open");
      }
      // Session closed before the hash landed → discard, never an orphan row.
      if (!this.sessions.has(sessionId)) {
        return this.revisions.prepare({
          sessionId,
          artifactRef,
          contentHash: stable.hash,
          origin: "external"
        });
      }
      const revision = this.revisions.prepare({
        sessionId,
        artifactRef,
        contentHash: stable.hash,
        origin: "external"
      });
      this.repos.insertRevision(revision);
      const live = this.sessions.get(sessionId);
      if (live && !live.session.committedRevision.contentHash) {
        this.updateSession(sessionId, (s) => {
          s.committedRevision = revision;
        });
        this.setReadiness(sessionId, { readiness: "ready", strongHash: stable.hash, probedAt: Date.now() });
      }
      return revision;
    })();
    this.strongIdentity.set(sessionId, strong);
    // A failed background hash surfaces on the next ensureStrongIdentity call.
    strong.catch(() => undefined);

    await this.events.emit(sessionId, 1, "session.opened", {
      documentId,
      artifactRef,
      revisionId: placeholder.revisionId
    });
    return session;
  }

  get(sessionId: SessionId): DocumentSession | undefined {
    return this.sessions.get(sessionId)?.session;
  }

  require(sessionId: SessionId): DocumentSession {
    const entry = this.sessions.get(sessionId);
    if (!entry) throw new OfficeRuntimeError("session-closed", `unknown session: ${sessionId}`);
    return entry.session;
  }

  actor(sessionId: SessionId): SessionActor {
    return this.requireEntry(sessionId).actor;
  }

  /**
   * §52/P0-6: writer paths (edit promotion, agent task, save, accept) wait
   * for the background strong hash here; interactive open never does.
   */
  async ensureStrongIdentity(sessionId: SessionId): Promise<CommittedRevision> {
    const pending = this.strongIdentity.get(sessionId);
    if (pending) return pending;
    const session = this.require(sessionId);
    if (session.committedRevision.contentHash) return session.committedRevision;
    throw new OfficeRuntimeError("recovery-required", `session ${sessionId} has no strong identity`);
  }

  readiness(sessionId: SessionId): WriteReadinessState {
    return this.requireEntry(sessionId).readiness;
  }

  setReadiness(sessionId: SessionId, readiness: WriteReadinessState): void {
    this.requireEntry(sessionId).readiness = readiness;
  }

  updateSession(sessionId: SessionId, mutate: (session: DocumentSession) => void): void {
    const entry = this.requireEntry(sessionId);
    mutate(entry.session);
    this.repos.upsertSession({
      sessionId,
      documentId: entry.session.documentId,
      artifactRef: entry.session.artifactRef,
      format: entry.session.format,
      backend: entry.session.backend,
      lifecycle: entry.session.lifecycle,
      epoch: entry.session.sessionEpoch,
      createdAt: entry.session.createdAt,
      closedAt: entry.session.lifecycle === "closed" ? Date.now() : null
    });
  }

  /** Bump epoch after recovery/rebind so in-flight results become stale (§43). */
  async bumpEpoch(sessionId: SessionId): Promise<number> {
    const entry = this.requireEntry(sessionId);
    entry.session.sessionEpoch += 1;
    this.repos.upsertSession({
      sessionId,
      documentId: entry.session.documentId,
      artifactRef: entry.session.artifactRef,
      format: entry.session.format,
      backend: entry.session.backend,
      lifecycle: entry.session.lifecycle,
      epoch: entry.session.sessionEpoch,
      createdAt: entry.session.createdAt
    });
    await this.events.emit(sessionId, entry.session.sessionEpoch, "session.epoch-bumped", {
      epoch: entry.session.sessionEpoch
    });
    return entry.session.sessionEpoch;
  }

  async close(sessionId: SessionId): Promise<void> {
    const entry = this.sessions.get(sessionId);
    if (!entry) return;
    entry.session.lifecycle = "closing";
    await entry.actor
      .enqueue(async () => {
        entry.session.lifecycle = "closed";
        this.sessions.delete(sessionId);
        this.repos.upsertSession({
          sessionId,
          documentId: entry.session.documentId,
          artifactRef: entry.session.artifactRef,
          format: entry.session.format,
          backend: entry.session.backend,
          lifecycle: "closed",
          epoch: entry.session.sessionEpoch,
          createdAt: entry.session.createdAt,
          closedAt: Date.now()
        });
        await this.events.emit(sessionId, entry.session.sessionEpoch, "session.closed", {});
      })
      .catch(() => undefined);
    this.strongIdentity.delete(sessionId);
  }

  list(): DocumentSession[] {
    return [...this.sessions.values()].map((e) => e.session);
  }

  /** Rehydrate sessions found in the DB after a crash (recovery-required lifecycle). */
  rehydrateFromDb(): number {
    let count = 0;
    for (const row of this.repos.listOpenSessions()) {
      if (this.sessions.has(row.session_id)) continue;
      const revision = this.revisions.latest(row.session_id);
      if (!revision) continue;
      const session: DocumentSession = {
        sessionId: row.session_id,
        documentId: row.document_id,
        format: row.format as OfficeFormat,
        backend: row.backend as DocumentSession["backend"],
        lifecycle: "recovery-required",
        sessionEpoch: row.epoch + 1,
        committedRevision: revision,
        artifactRef: row.artifact_ref,
        createdAt: row.created_at
      };
      this.sessions.set(row.session_id, { session, actor: new SessionActor(), readiness: { readiness: "cold" } });
      this.repos.upsertSession({
        sessionId: row.session_id,
        documentId: row.document_id,
        artifactRef: row.artifact_ref,
        format: session.format,
        backend: session.backend,
        lifecycle: "recovery-required",
        epoch: session.sessionEpoch,
        createdAt: row.created_at
      });
      count++;
    }
    return count;
  }

  private async resolveOrCreateDocumentId(artifactRef: ArtifactRef): Promise<DocumentId> {
    // DocumentId is per source artifact path; reopening the same file shares identity.
    const path = this.store.resolvePath(artifactRef);
    for (const row of this.repos.loadArtifacts()) {
      if (row.path === path) {
        const existing = this.repos
          .listOpenSessions()
          .find((s) => this.store.tryResolvePath(s.artifact_ref) === path);
        if (existing) return existing.document_id;
      }
    }
    return newDocumentId();
  }

  private requireEntry(sessionId: SessionId): { session: DocumentSession; actor: SessionActor; readiness: WriteReadinessState } {
    const entry = this.sessions.get(sessionId);
    if (!entry) throw new OfficeRuntimeError("session-closed", `unknown session: ${sessionId}`);
    return entry;
  }
}
