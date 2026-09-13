/**
 * RecoveryService (§77, §141): resolves crash-interrupted commits. Truth
 * priority: filesystem hash facts > CommitJournal > session metadata. A
 * journal row in `committing` alone never guesses — the source file's actual
 * hash decides finalize vs rollback vs conflict (INV-15).
 *
 * P0 hardening (round 6):
 *  - SOURCE_REPLACED forward-finalize uses the SAME atomic transaction as the
 *    normal committer (prepare + finalizeCommitAtomically) — no crash window
 *    between revision insert and journal flip.
 *  - Source missing (hash undefined) is a CONFLICT that PRESERVES the temp
 *    file — never deletes the only recoverable candidate copy.
 *  - source-unregistrable keeps the journal UNRESOLVED (recovery-blocked),
 *    never marks aborted when the physical commit already replaced the source.
 *
 * Round 7 (schema v3): commit idempotency is EXACT — revisions.commit_id is
 * UNIQUE, so "this commit already landed" is identity-proven. A→B→A content
 * returns never collide with history; every interrupted commit lands its own
 * revision with a continuous sequence.
 */

import type { RecoveryOutcome } from "../../contracts/revision.js";
import { sha256File } from "../../support/fsx.js";
import type { RuntimeRepositories } from "../persistence/repositories.js";
import type { RevisionLog } from "../revisions/revision-log.js";
import type { DurableEventBus } from "../sessions/event-bus.js";
import type { AtomicFileCommitter } from "../commit/atomic-file-committer.js";
import type { ArtifactStore } from "../../artifact/store/artifact-store.js";
import type { CommitJournalRecord } from "../../contracts/revision.js";

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

    // ── P0-2: source missing is ALWAYS a conflict that preserves evidence ──
    // The temp file may be the ONLY recoverable candidate copy. Never delete.
    if (sourceHash === undefined) {
      await this.emit(record.sessionId, {
        commitId: record.commitId,
        resolution: "conflict",
        reason: "source-missing-preserve-temp"
      });
      return { commitId: record.commitId, resolution: "conflict", reason: "source-missing-preserve-temp" };
    }

    if (record.phase === "prepared" || record.phase === "temp-ready") {
      // Replace never happened: source at base hash → safe rollback.
      if (sourceHash === record.sourceHashBefore) {
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
      // Replace happened; hash facts decide.
      if (sourceHash === record.candidateHash) {
        // ── P0-1: use the SAME atomic finalization as the normal committer.
        // revision.prepare → finalizeCommitAtomically → journal FINALIZED in
        // ONE transaction. No crash window between insert and phase flip.
        let artifactRef: string;
        try {
          artifactRef = await this.resolveArtifactRef(record.sourcePath);
        } catch {
          // P1-high: the physical commit ALREADY replaced the source — this
          // is recovery-blocked, NOT aborted. Keep the journal unresolved so
          // the next recovery (or manual intervention) can finalize once the
          // artifact becomes registrable.
          await this.emit(record.sessionId, {
            commitId: record.commitId,
            resolution: "conflict",
            reason: "source-unregistrable-recovery-blocked"
          });
          return {
            commitId: record.commitId,
            resolution: "conflict",
            reason: "source-unregistrable-recovery-blocked"
          };
        }

        const journalRecord = this.repos.getJournal(record.commitId);
        if (!journalRecord) {
          return { commitId: record.commitId, resolution: "conflict", reason: "journal-vanished" };
        }

        // v3 exact-commit idempotency: "has THIS commit landed a revision?"
        // is decided by commit_id identity — NEVER by content hash. A→B→A
        // same-origin commits each land their own revision; a content-hash
        // match against history would swallow the newest revision and lose
        // the sequence.
        const existing = this.revisions.getRevisionByCommitId(record.commitId);
        if (!existing) {
          const revision = this.revisions.prepare({
            sessionId: record.sessionId,
            artifactRef,
            contentHash: record.candidateHash,
            origin: record.origin as "human" | "agent" | "external",
            commitId: record.commitId
          });
          this.repos.finalizeCommitAtomically(revision, journalRecord);
        } else {
          // The revision for THIS exact commit exists but the journal never
          // flipped (injected/legacy state) — atomic identity-proven flip.
          this.repos.markJournalFinalized(record.commitId);
        }

        await this.emit(record.sessionId, {
          commitId: record.commitId,
          resolution: "finalized",
          contentHash: record.candidateHash
        });
        return { commitId: record.commitId, resolution: "finalized", contentHash: record.candidateHash };
      }
      if (sourceHash === record.sourceHashBefore) {
        // Journal says replaced but source is at base → rollback.
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
   * P0 FAIL CLOSED: if the source cannot be registered, recovery reports
   * recovery-blocked — never fabricates an unresolvable ArtifactRef.
   */
  private async resolveArtifactRef(sourcePath: string): Promise<string> {
    const existing = this.store.tryResolveRefByPath(sourcePath);
    if (existing) return existing;
    const format = sourcePath.split(".").pop() as "docx" | "xlsx" | "pptx";
    return await this.store.register(sourcePath, { format });
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
