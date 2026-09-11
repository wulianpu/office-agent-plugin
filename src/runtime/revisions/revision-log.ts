/**
 * RevisionLog (§36): committed revisions are the MVCC read anchors. Human
 * saves, agent accepts and external imports all finalize here.
 */

import type { ArtifactRef, SessionId } from "../../contracts/ids.js";
import type { CommittedRevision, RevisionOrigin } from "../../contracts/revision.js";
import { newRevisionId } from "../../support/ids.js";
import type { RuntimeRepositories } from "../persistence/repositories.js";

export class RevisionLog {
  constructor(private readonly repos: RuntimeRepositories) {}

  nextSequence(sessionId: SessionId): number {
    const latest = this.repos.latestRevision(sessionId);
    return (latest?.sequence ?? 0) + 1;
  }

  /** Reserve id+sequence WITHOUT inserting — the insert happens atomically
   *  with the journal flip (P0-A, repos.finalizeCommitAtomically). */
  prepare(params: {
    sessionId: SessionId;
    artifactRef: ArtifactRef;
    contentHash: string;
    origin: RevisionOrigin;
  }): CommittedRevision {
    return {
      revisionId: newRevisionId(),
      sessionId: params.sessionId,
      sequence: this.nextSequence(params.sessionId),
      artifactRef: params.artifactRef,
      contentHash: params.contentHash,
      origin: params.origin,
      createdAt: Date.now()
    };
  }

  commit(params: {
    sessionId: SessionId;
    artifactRef: ArtifactRef;
    contentHash: string;
    origin: RevisionOrigin;
  }): CommittedRevision {
    const revision: CommittedRevision = {
      revisionId: newRevisionId(),
      sessionId: params.sessionId,
      sequence: this.nextSequence(params.sessionId),
      artifactRef: params.artifactRef,
      contentHash: params.contentHash,
      origin: params.origin,
      createdAt: Date.now()
    };
    this.repos.insertRevision(revision);
    return revision;
  }

  latest(sessionId: SessionId): CommittedRevision | undefined {
    return this.repos.latestRevision(sessionId);
  }

  list(sessionId: SessionId): CommittedRevision[] {
    return this.repos.listRevisions(sessionId);
  }

  get(revisionId: string): CommittedRevision | undefined {
    return this.repos.getRevision(revisionId);
  }
}
