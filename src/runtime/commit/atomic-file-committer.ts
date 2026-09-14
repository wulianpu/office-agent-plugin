/**
 * AtomicFileCommitter (§74–§76): every source replacement — human save, agent
 * accept, external import — flows through here. The journal records
 * PREPARED → TEMP_READY → SOURCE_REPLACED → FINALIZED so crash recovery can
 * reconstruct truth from filesystem hashes (§77, INV-15).
 *
 * The committer operates on physical paths; the runtime service resolves
 * ArtifactRefs via the ArtifactStore before invoking it.
 */

import { copyFile, rename, rm } from "node:fs/promises";
import type { CommitJournalRecord, CommitPhase, CommitResult, RevisionOrigin } from "../../contracts/revision.js";
import { OfficeRuntimeError } from "../../contracts/document.js";
import { newCommitId } from "../../support/ids.js";
import { fsyncFile, pathExists, removeQuiet, sha256File } from "../../support/fsx.js";
import type { RuntimeRepositories } from "../persistence/repositories.js";
import type { RevisionLog } from "../revisions/revision-log.js";
import type { DurableEventBus } from "../sessions/event-bus.js";
import type { SelfWriteGuardRegistry } from "./self-write-guard.js";

export interface CommitPathsRequest {
  sessionId: string;
  sessionEpoch?: number;
  candidateId: string;
  sourcePath: string;
  candidatePath: string;
  /** ArtifactRef the committed revision records for the source (required —
   *  the revision row must never persist an empty ref). */
  sessionArtifactRef: string;
  /** Hash the source must still have at replace time (INV-10). */
  expectedSourceHash: string;
  candidateHash: string;
  origin: RevisionOrigin;
}

export class AtomicFileCommitter {
  /** Optional lock releaser invoked when Windows rename hits EPERM/EBUSY. */
  lockReleaser?: (path: string) => Promise<void>;
  /**
   * §141 fault injection (test-only): invoked after each journal phase lands.
   * A hook that kills the process here reproduces a REAL crash at a commit
   * phase boundary (stronger than fabricating journal rows by hand).
   */
  faultHook?: (phase: CommitPhase, commitId: string) => void;

  constructor(
    private readonly repos: RuntimeRepositories,
    private readonly revisions: RevisionLog,
    private readonly events: DurableEventBus,
    private readonly selfWrites: SelfWriteGuardRegistry
  ) {}

  /**
   * Windows replace with lock-resilient retry: engine daemons or AV scanners
   * holding the destination clear within a few hundred ms; a releaser hook
   * (officecli close) resolves persistent engine residents.
   */
  private async replaceAtomically(tempPath: string, sourcePath: string): Promise<void> {
    const attempts = [
      { wait: 0 },
      { wait: 120 },
      { wait: 300 },
      { wait: 600, release: true },
      { wait: 1200, release: true }
    ];
    let lastError: unknown;
    for (const attempt of attempts) {
      if (attempt.wait > 0) {
        await new Promise((resolve) => setTimeout(resolve, attempt.wait));
      }
      if (attempt.release && this.lockReleaser) {
        await this.lockReleaser(sourcePath).catch(() => undefined);
      }
      try {
        await rename(tempPath, sourcePath);
        return;
      } catch (error) {
        lastError = error;
        const code = (error as NodeJS.ErrnoException).code;
        if (code !== "EPERM" && code !== "EBUSY" && code !== "EACCES") throw error;
      }
    }
    throw lastError;
  }

  /**
   * verify expected source hash → prepare sibling temp → fsync → replace →
   * rehash → finalize revision. rename() replaces the destination file
   * atomically on POSIX and via MOVEFILE_REPLACE_EXISTING semantics on Windows.
   */
  async commit(request: CommitPathsRequest): Promise<CommitResult> {
    const commitId = newCommitId();
    const tempPath = `${request.sourcePath}.commit-${commitId}`;
    const record: CommitJournalRecord = {
      commitId,
      sessionId: request.sessionId,
      candidateId: request.candidateId,
      sourcePath: request.sourcePath,
      tempPath,
      sourceHashBefore: request.expectedSourceHash,
      candidateHash: request.candidateHash,
      phase: "prepared",
      origin: request.origin,
      createdAt: Date.now(),
      updatedAt: Date.now()
    };
    await this.journal(record, "prepared");

    // INV-10: the source must still be at the base hash.
    const currentHash = await sha256File(request.sourcePath).catch(() => undefined);
    if (currentHash === undefined) {
      await this.journal(record, "aborted");
      throw new OfficeRuntimeError("artifact-missing", `source vanished before commit: ${request.sourcePath}`);
    }
    if (currentHash !== request.expectedSourceHash) {
      await this.journal(record, "aborted");
      throw new OfficeRuntimeError(
        "source-mutated",
        `source changed before commit: expected ${request.expectedSourceHash}, found ${currentHash}`
      );
    }

    await copyFile(request.candidatePath, tempPath);
    await fsyncFile(tempPath);
    await this.journal(record, "temp-ready");

    // ── P0 candidate seal (#5): the bytes about to be renamed must BE the
    // verification-bound hash, verified BEFORE the destructive replace. The
    // sibling temp is this commit's own fsynced copy, so its hash is stable;
    // a mismatch (staging changed after verification, wrong call path) must
    // fail here — previously the unverified bytes were renamed onto the
    // source FIRST and only discovered by the post-replace rehash, after
    // the old source content was already lost.
    const tempHash = await sha256File(tempPath);
    if (tempHash !== request.candidateHash) {
      await this.cleanupTemp(record);
      await this.journal(record, "aborted");
      throw new OfficeRuntimeError(
        "candidate-hash-mismatch",
        `candidate seal failed: temp holds ${tempHash}, verification bound ${request.candidateHash}`
      );
    }

    // ── P0 source seal (#5): the authoritative INV-10 check moves to the
    // last gate before the replace. The early check above can pass and an
    // external Word/WPS/sync save can still land inside the copy + fsync +
    // journal window; that external version must survive untouched instead
    // of being silently overwritten by the rename. (A residual OS-level
    // race between this check and the rename syscall remains — no
    // cross-platform hash-CAS-rename primitive exists — but the large
    // window is closed.)
    const finalSourceHash = await sha256File(request.sourcePath).catch(() => undefined);
    if (finalSourceHash === undefined) {
      await this.cleanupTemp(record);
      await this.journal(record, "aborted");
      throw new OfficeRuntimeError(
        "artifact-missing",
        `source vanished before replace: ${request.sourcePath}`
      );
    }
    if (finalSourceHash !== request.expectedSourceHash) {
      await this.cleanupTemp(record);
      await this.journal(record, "aborted");
      throw new OfficeRuntimeError(
        "source-mutated",
        `source changed during commit (pre-replace seal): expected ${request.expectedSourceHash}, found ${finalSourceHash} — external bytes preserved`
      );
    }

    await this.replaceAtomically(tempPath, request.sourcePath);
    await fsyncFile(request.sourcePath);
    // P1 hardening: POSIX parent-dir fsync makes the rename durable across
    // OS crash / sudden power loss (not just process crash). Windows
    // doesn't support directory fsync — handled inside the helper.
    const { fsyncParentDir } = await import("../../support/fsx.js");
    await fsyncParentDir(request.sourcePath);
    await this.journal(record, "source-replaced");

    // Final durability/integrity assertion — NOT the first line of defense
    // (the pre-replace seals are). A mismatch here means a filesystem
    // anomaly or a last-instant external write raced the rename: surface it
    // as recovery-required (the journal sits at SOURCE_REPLACED with hash
    // facts recorded; Recovery resolves the conflict), never as a plain
    // retryable io-error.
    const replacedHash = await sha256File(request.sourcePath);
    if (replacedHash !== request.candidateHash) {
      throw new OfficeRuntimeError(
        "recovery-required",
        `post-replace hash mismatch: ${replacedHash} != ${request.candidateHash} — recovery required`
      );
    }

    // P0-A: build the revision, then land revision-insert + journal
    // FINALIZED in ONE transaction (see finalizeCommitAtomically). The
    // revision carries this commit's identity (v3) — recovery decides
    // "already landed?" by exact commit_id, never by content hash.
    const revision = this.revisions.prepare({
      sessionId: request.sessionId,
      artifactRef: request.sessionArtifactRef,
      contentHash: replacedHash,
      origin: request.origin,
      commitId
    });
    this.repos.finalizeCommitAtomically(revision, record);

    this.selfWrites.register({
      commitId,
      expectedHash: replacedHash,
      expiresAt: Date.now() + 10_000,
      sourcePath: request.sourcePath
    });
    await this.events.emit(request.sessionId, request.sessionEpoch ?? 0, "revision.committed", {
      revisionId: revision.revisionId,
      sequence: revision.sequence,
      origin: revision.origin,
      contentHash: replacedHash,
      commitId
    });
    return { commitId, newRevision: revision, journalOutcome: "finalized" };
  }

  private async journal(record: CommitJournalRecord, phase: CommitPhase): Promise<void> {
    const updated = { ...record, phase, updatedAt: Date.now() };
    this.repos.upsertJournal(updated);
    // §141: fire AFTER the phase is durable — a hook that kills the process
    // here leaves the journal exactly at this phase boundary.
    this.faultHook?.(phase, record.commitId);
    await this.events
      .emit(record.sessionId, 0, "commit.journal-updated", { commitId: record.commitId, phase })
      .catch(() => undefined);
  }

  /** Cleanup helper for aborted/rolled-back commits. */
  async cleanupTemp(record: CommitJournalRecord): Promise<void> {
    if (await pathExists(record.tempPath)) await removeQuiet(record.tempPath);
    await rm(`${record.sourcePath}.commit-${record.commitId}`, { force: true });
  }
}
