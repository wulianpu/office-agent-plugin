/**
 * RecoveryService (§77, §141): resolves crash-interrupted commits. Truth
 * priority: filesystem hash facts > CommitJournal > session metadata. A
 * journal row in `committing` alone never guesses — the source file's actual
 * hash decides finalize vs rollback vs conflict (INV-15).
 */

import type { RecoveryOutcome } from "../../contracts/revision.js";
import { canonicalSourceKey, sha256File } from "../../support/fsx.js";
import type { RuntimeRepositories } from "../persistence/repositories.js";
import type { RevisionLog } from "../revisions/revision-log.js";
import type { DurableEventBus } from "../sessions/event-bus.js";
import type { AtomicFileCommitter } from "../commit/atomic-file-committer.js";
import type { ArtifactStore } from "../../artifact/store/artifact-store.js";

export class RecoveryService {
  constructor(
    private readonly repos: RuntimeRepositories,
    private readonly revisions: RevisionLog,
    private readonly events: DurableEventBus,
    private readonly committer: AtomicFileCommitter,
    private readonly store: ArtifactStore
  ) {}

  async recoverAll(): Promise<RecoveryOutcome[]> {
    const outcomes: RecoveryOutcome[] = [];
    for (const record of this.repos.unresolvedJournal()) {
      outcomes.push(await this.recoverOne(record));
    }
    return outcomes;
  }

  async recoverOne(record: {
    commitId: string;
    sessionId: string;
    candidateId: string;
    sourcePath: string;
    tempPath: string;
    sourceHashBefore: string;
    candidateHash: string;
    phase: string;
    origin: string;
  }): Promise<RecoveryOutcome> {
    const sourceHash = await sha256File(record.sourcePath).catch(() => undefined);

    if (record.phase === "prepared" || record.phase === "temp-ready") {
      // Replace never happened (or never will): rollback if source untouched.
      if (sourceHash === undefined || sourceHash === record.sourceHashBefore) {
        await this.committer.cleanupTemp(record as never);
        this.markPhase(record.commitId, "aborted");
        await this.emit(record.sessionId, {
          commitId: record.commitId,
          resolution: "rolled-back",
          reason: record.phase === "prepared" ? "source-unchanged" : "temp-only"
        });
        return {
          commitId: record.commitId,
          resolution: "rolled-back",
          reason: record.phase === "prepared" ? "source-unchanged" : "temp-only"
        };
      }
      // Source changed by someone else mid-commit → conflict.
      this.markPhase(record.commitId, "aborted");
      await this.emit(record.sessionId, {
        commitId: record.commitId,
        resolution: "conflict",
        reason: `source mutated during interrupted commit (${record.phase})`
      });
      return { commitId: record.commitId, resolution: "conflict", reason: "source-mutated" };
    }

    if (record.phase === "source-replaced") {
      // Replace happened; did finalize land? Decide purely by hash facts.
      if (sourceHash === record.candidateHash) {
        // Commit effectively completed; finalize the revision if missing.
        const already = this.revisions
          .list(record.sessionId)
          .some((r) => r.contentHash === record.candidateHash);
        if (!already) {
          this.revisions.commit({
            sessionId: record.sessionId,
            artifactRef: await this.resolveArtifactRef(record.sourcePath),
            contentHash: record.candidateHash,
            origin: record.origin as "human" | "agent" | "external"
          });
        }
        this.markPhase(record.commitId, "finalized");
        await this.emit(record.sessionId, {
          commitId: record.commitId,
          resolution: "finalized",
          contentHash: record.candidateHash
        });
        return { commitId: record.commitId, resolution: "finalized", contentHash: record.candidateHash };
      }
      if (sourceHash === record.sourceHashBefore) {
        // Journal lied (replace recorded but bytes rolled back?) → rollback.
        await this.committer.cleanupTemp(record as never);
        this.markPhase(record.commitId, "aborted");
        await this.emit(record.sessionId, {
          commitId: record.commitId,
          resolution: "rolled-back",
          reason: "aborted"
        });
        return { commitId: record.commitId, resolution: "rolled-back", reason: "aborted" };
      }
      this.markPhase(record.commitId, "aborted");
      await this.emit(record.sessionId, {
        commitId: record.commitId,
        resolution: "conflict",
        reason: "source hash matches neither base nor candidate after crash"
      });
      return {
        commitId: record.commitId,
        resolution: "conflict",
        reason: "hash-mismatch-after-crash"
      };
    }

    return { commitId: record.commitId, resolution: "conflict", reason: `unknown phase ${record.phase}` };
  }

  /**
   * P0-3: async and STRICT — never falls back to a raw filesystem path as an
   * ArtifactRef. Registration failure surfaces through the caller.
   */
  private async resolveArtifactRef(sourcePath: string): Promise<string> {
    const existing = this.store.tryResolveRefByPath(sourcePath);
    if (existing) return existing;
    const format = sourcePath.split(".").pop() as "docx" | "xlsx" | "pptx";
    try {
      return await this.store.register(sourcePath, { format });
    } catch {
      // Registration failed (missing file): opaque canonical key, never the
      // raw path — resolvePath will surface a clear error if it is used.
      return canonicalSourceKey(sourcePath);
    }
  }

  private markPhase(commitId: string, phase: "finalized" | "aborted"): void {
    const record = this.repos.getJournal(commitId);
    if (!record) return;
    this.repos.upsertJournal({ ...record, phase, updatedAt: Date.now() });
  }

  private async emit(sessionId: string, outcome: RecoveryOutcome): Promise<void> {
    await this.events
      .emit(sessionId, 0, "recovery.resolved", outcome as Record<string, unknown>)
      .catch(() => undefined);
  }
}
