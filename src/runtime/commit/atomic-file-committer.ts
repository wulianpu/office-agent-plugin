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
  /** ArtifactRef the committed revision records for the source. */
  sessionArtifactRef?: string;
  /** Hash the source must still have at replace time (INV-10). */
  expectedSourceHash: string;
  candidateHash: string;
  origin: RevisionOrigin;
}

export class AtomicFileCommitter {
  /** Optional lock releaser invoked when Windows rename hits EPERM/EBUSY. */
  lockReleaser?: (path: string) => Promise<void>;

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

    await this.replaceAtomically(tempPath, request.sourcePath);
    await fsyncFile(request.sourcePath);
    await this.journal(record, "source-replaced");

    // Rehash the replaced source; must equal the candidate bytes.
    const replacedHash = await sha256File(request.sourcePath);
    if (replacedHash !== request.candidateHash) {
      throw new OfficeRuntimeError(
        "io-error",
        `post-replace hash mismatch: ${replacedHash} != ${request.candidateHash}`
      );
    }

    await this.journal(record, "finalized");

    const revision = this.revisions.commit({
      sessionId: request.sessionId,
      artifactRef: "", // service层填充 artifactRef（见 OfficeRuntimeService.acceptCandidate）
      contentHash: replacedHash,
      origin: request.origin
    });
    const revisionWithRef = { ...revision, artifactRef: request.sessionArtifactRef ?? revision.artifactRef };

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
    return { commitId, newRevision: revisionWithRef, journalOutcome: "finalized" };
  }

  private async journal(record: CommitJournalRecord, phase: CommitPhase): Promise<void> {
    const updated = { ...record, phase, updatedAt: Date.now() };
    this.repos.upsertJournal(updated);
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
