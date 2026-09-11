/**
 * Typed repositories over RuntimeDatabase. All bigint values cross the SQLite
 * boundary as TEXT. Each method runs in a short transaction or a single
 * statement — never a long-lived read transaction (§119).
 */

import type {
  ArtifactRef,
  CandidateId,
  LeaseId,
  OfficeFormat,
  RevisionId,
  SessionEpoch,
  SessionId
} from "../../contracts/ids.js";
import type { CommitJournalRecord, CommitPhase, RevisionOrigin } from "../../contracts/revision.js";
import type { CandidateRevision, CandidateState } from "../../contracts/candidate.js";
import type { CommittedRevision } from "../../contracts/revision.js";
import type { WriterLease } from "../../contracts/lease.js";
import type { SessionEvent, SessionEventType } from "../../contracts/events.js";
import type { MutationReceipt } from "../../contracts/capabilities.js";
import type { StoredArtifact } from "../../artifact/store/artifact-store.js";
import type { ArtifactStorePersistence } from "../../artifact/store/store-persistence.js";
import type { ScanPersistence } from "../../artifact/scanner/scanner.js";
import type { RuntimeDatabase } from "./database.js";

interface SessionRow {
  session_id: string;
  document_id: string;
  artifact_ref: string;
  format: string;
  backend: string;
  lifecycle: string;
  epoch: number;
  created_at: number;
  closed_at: number | null;
}

interface RevisionRow {
  revision_id: string;
  session_id: string;
  sequence: number;
  artifact_ref: string;
  content_hash: string;
  origin: string;
  created_at: number;
}

interface CandidateRow {
  candidate_id: string;
  session_id: string;
  base_revision_id: string;
  base_hash: string;
  artifact_ref: string;
  current_hash: string | null;
  state: string;
  created_by: string;
  verification: string | null;
  failure_reason: string | null;
  created_at: number;
  updated_at: number;
}

interface LeaseRow {
  lease_id: string;
  session_id: string;
  owner: string;
  backend: string;
  base_revision_id: string;
  fencing_token: string;
  epoch: number;
  acquired_at: number;
  released_at: number | null;
}

interface EventRow {
  event_id: string;
  session_id: string;
  sequence: string;
  epoch: number;
  type: string;
  payload: string;
  created_at: number;
}

interface JournalRow {
  commit_id: string;
  session_id: string;
  candidate_id: string;
  source_path: string;
  temp_path: string;
  source_hash_before: string;
  candidate_hash: string;
  phase: string;
  origin: string;
  created_at: number;
  updated_at: number;
}

export class RuntimeRepositories
  implements ArtifactStorePersistence, ScanPersistence
{
  constructor(private readonly rtdb: RuntimeDatabase) {}

  private get db() {
    return this.rtdb.db;
  }

  // ---- artifacts (ArtifactStorePersistence) ----

  saveArtifact(row: StoredArtifact): Promise<void> {
    this.db
      .prepare(
        "INSERT OR REPLACE INTO artifacts (ref, path, kind, format, created_at) VALUES (?, ?, ?, ?, ?)"
      )
      .run(row.ref, row.path, row.kind, row.format, Date.now());
    return Promise.resolve();
  }

  deleteArtifact(ref: ArtifactRef): Promise<void> {
    this.db.prepare("DELETE FROM artifacts WHERE ref = ?").run(ref);
    return Promise.resolve();
  }

  loadArtifacts(): StoredArtifact[] {
    return (
      this.db.prepare("SELECT ref, path, kind, format FROM artifacts").all() as Array<{
        ref: string;
        path: string;
        kind: string;
        format: string;
      }>
    ).map((r) => ({
      ref: r.ref,
      path: r.path,
      kind: r.kind as "source" | "staging",
      format: r.format as OfficeFormat
    }));
  }

  // ---- scans (ScanPersistence) ----

  saveScan(row: {
    contentHash: string;
    size: string;
    manifest: string;
    corruptEntries: string[];
    scannedAt: number;
  }): Promise<void> {
    this.db
      .prepare(
        "INSERT OR REPLACE INTO scans (content_hash, size, manifest, corrupt_entries, scanned_at) VALUES (?, ?, ?, ?, ?)"
      )
      .run(row.contentHash, row.size, row.manifest, JSON.stringify(row.corruptEntries), row.scannedAt);
    return Promise.resolve();
  }

  // ---- sessions ----

  upsertSession(row: {
    sessionId: SessionId;
    documentId: string;
    artifactRef: ArtifactRef;
    format: OfficeFormat;
    backend: string;
    lifecycle: string;
    epoch: SessionEpoch;
    createdAt: number;
    closedAt?: number | null;
  }): void {
    this.db
      .prepare(
        `INSERT INTO sessions (session_id, document_id, artifact_ref, format, backend, lifecycle, epoch, created_at, closed_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
         ON CONFLICT(session_id) DO UPDATE SET lifecycle = excluded.lifecycle, epoch = excluded.epoch, closed_at = excluded.closed_at`
      )
      .run(
        row.sessionId,
        row.documentId,
        row.artifactRef,
        row.format,
        row.backend,
        row.lifecycle,
        row.epoch,
        row.createdAt,
        row.closedAt ?? null
      );
  }

  getSession(sessionId: SessionId): SessionRow | undefined {
    return this.db.prepare("SELECT * FROM sessions WHERE session_id = ?").get(sessionId) as
      | SessionRow
      | undefined;
  }

  listOpenSessions(): SessionRow[] {
    return this.db
      .prepare("SELECT * FROM sessions WHERE lifecycle NOT IN ('closed') ORDER BY created_at")
      .all() as unknown as SessionRow[];
  }

  // ---- revisions ----

  insertRevision(rev: CommittedRevision): void {
    this.db
      .prepare(
        `INSERT INTO revisions (revision_id, session_id, sequence, artifact_ref, content_hash, origin, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?)`
      )
      .run(
        rev.revisionId,
        rev.sessionId,
        rev.sequence,
        rev.artifactRef,
        rev.contentHash,
        rev.origin,
        rev.createdAt
      );
  }

  getRevision(revisionId: RevisionId): CommittedRevision | undefined {
    const row = this.db.prepare("SELECT * FROM revisions WHERE revision_id = ?").get(revisionId) as
      | RevisionRow
      | undefined;
    return row ? revisionFromRow(row) : undefined;
  }

  latestRevision(sessionId: SessionId): CommittedRevision | undefined {
    const row = this.db
      .prepare("SELECT * FROM revisions WHERE session_id = ? ORDER BY sequence DESC LIMIT 1")
      .get(sessionId) as RevisionRow | undefined;
    return row ? revisionFromRow(row) : undefined;
  }

  listRevisions(sessionId: SessionId): CommittedRevision[] {
    return (
      this.db
        .prepare("SELECT * FROM revisions WHERE session_id = ? ORDER BY sequence")
        .all(sessionId) as unknown as RevisionRow[]
    ).map(revisionFromRow);
  }

  // ---- candidates ----

  upsertCandidate(c: CandidateRevision): void {
    this.db
      .prepare(
        `INSERT INTO candidates (candidate_id, session_id, base_revision_id, base_hash, artifact_ref, current_hash, state, created_by, verification, failure_reason, created_at, updated_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
         ON CONFLICT(candidate_id) DO UPDATE SET
           current_hash = excluded.current_hash,
           state = excluded.state,
           verification = excluded.verification,
           failure_reason = excluded.failure_reason,
           updated_at = excluded.updated_at`
      )
      .run(
        c.candidateId,
        c.sessionId,
        c.baseRevisionId,
        c.baseHash,
        c.artifactRef,
        c.currentHash ?? null,
        c.state,
        c.createdBy,
        c.verification ? JSON.stringify(c.verification) : null,
        c.failureReason ?? null,
        c.createdAt,
        c.updatedAt
      );
  }

  getCandidate(candidateId: CandidateId): CandidateRevision | undefined {
    const row = this.db
      .prepare("SELECT * FROM candidates WHERE candidate_id = ?")
      .get(candidateId) as CandidateRow | undefined;
    return row ? candidateFromRow(row) : undefined;
  }

  /** Candidate lifecycle ends at accept — the revision carries the truth. */
  deleteCandidate(candidateId: CandidateId): void {
    this.db.prepare("DELETE FROM candidates WHERE candidate_id = ?").run(candidateId);
  }

  activeCandidateForSession(sessionId: SessionId): CandidateRevision | undefined {
    const row = this.db
      .prepare(
        "SELECT * FROM candidates WHERE session_id = ? AND state NOT IN ('failed') ORDER BY created_at DESC LIMIT 1"
      )
      .get(sessionId) as CandidateRow | undefined;
    return row ? candidateFromRow(row) : undefined;
  }

  // ---- leases ----

  insertLease(lease: WriterLease): void {
    this.db
      .prepare(
        `INSERT INTO leases (lease_id, session_id, owner, backend, base_revision_id, fencing_token, epoch, acquired_at, released_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULL)`
      )
      .run(
        lease.leaseId,
        lease.sessionId,
        lease.owner,
        lease.backend,
        lease.baseRevisionId,
        lease.fencingToken.toString(),
        lease.sessionEpoch,
        lease.acquiredAt
      );
  }

  releaseLease(leaseId: LeaseId): void {
    this.db
      .prepare("UPDATE leases SET released_at = ? WHERE lease_id = ?")
      .run(Date.now(), leaseId);
  }

  latestFencingToken(sessionId: SessionId): bigint {
    const row = this.db
      .prepare("SELECT MAX(CAST(fencing_token AS INTEGER)) AS max FROM leases WHERE session_id = ?")
      .get(sessionId) as { max: number | null };
    // CAST to INTEGER in SQLite is 64-bit safe; the driver returns number only
    // when within safe range, so parse the text form instead.
    const text = this.db
      .prepare(
        "SELECT fencing_token FROM leases WHERE session_id = ? ORDER BY CAST(fencing_token AS INTEGER) DESC LIMIT 1"
      )
      .get(sessionId) as { fencing_token: string } | undefined;
    if (text) return BigInt(text.fencing_token);
    return 0n;
  }

  // ---- events (SessionEventStore) ----

  appendEvent(event: SessionEvent): void {
    this.db
      .prepare(
        `INSERT INTO events (event_id, session_id, sequence, epoch, type, payload, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?)`
      )
      .run(
        event.eventId,
        event.sessionId,
        event.sequence.toString(),
        event.sessionEpoch,
        event.type,
        JSON.stringify(event.payload),
        event.createdAt
      );
  }

  readEventsSince(sessionId: SessionId, afterSequence: bigint, limit = 500): SessionEvent[] {
    return (
      this.db
        .prepare(
          "SELECT * FROM events WHERE session_id = ? AND CAST(sequence AS INTEGER) > ? ORDER BY CAST(sequence AS INTEGER) LIMIT ?"
        )
        .all(sessionId, Number(afterSequence), limit) as unknown as EventRow[]
    ).map((row) => ({
      eventId: row.event_id,
      sessionId: row.session_id,
      sequence: BigInt(row.sequence),
      sessionEpoch: row.epoch,
      type: row.type as SessionEventType,
      payload: JSON.parse(row.payload),
      createdAt: row.created_at
    }));
  }

  lastEventSequence(sessionId: SessionId): bigint {
    const row = this.db
      .prepare(
        "SELECT sequence FROM events WHERE session_id = ? ORDER BY CAST(sequence AS INTEGER) DESC LIMIT 1"
      )
      .get(sessionId) as { sequence: string } | undefined;
    return row ? BigInt(row.sequence) : 0n;
  }

  // ---- commit journal ----

  upsertJournal(record: CommitJournalRecord): void {
    this.db
      .prepare(
        `INSERT INTO commit_journal (commit_id, session_id, candidate_id, source_path, temp_path, source_hash_before, candidate_hash, phase, origin, created_at, updated_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
         ON CONFLICT(commit_id) DO UPDATE SET phase = excluded.phase, updated_at = excluded.updated_at`
      )
      .run(
        record.commitId,
        record.sessionId,
        record.candidateId,
        record.sourcePath,
        record.tempPath,
        record.sourceHashBefore,
        record.candidateHash,
        record.phase,
        record.origin,
        record.createdAt,
        record.updatedAt
      );
  }

  getJournal(commitId: string): CommitJournalRecord | undefined {
    const row = this.db
      .prepare("SELECT * FROM commit_journal WHERE commit_id = ?")
      .get(commitId) as JournalRow | undefined;
    return row ? journalFromRow(row) : undefined;
  }

  unresolvedJournal(): CommitJournalRecord[] {
    return (
      this.db
        .prepare(
          "SELECT * FROM commit_journal WHERE phase NOT IN ('finalized', 'aborted') ORDER BY created_at"
        )
        .all() as unknown as JournalRow[]
    ).map(journalFromRow);
  }

  listJournal(): CommitJournalRecord[] {
    return (this.db.prepare("SELECT * FROM commit_journal ORDER BY created_at").all() as unknown as JournalRow[]).map(
      journalFromRow
    );
  }

  // ---- idempotency (v2: candidate-scoped, §61) ----

  saveIdempotentReceipt(input: {
    candidateId: string;
    idempotencyKey: string;
    commandId: string;
    payloadDigest: string;
    receipt: MutationReceipt;
  }): void {
    this.db
      .prepare(
        `INSERT OR IGNORE INTO idempotency (candidate_id, idempotency_key, command_id, payload_digest, receipt, created_at)
         VALUES (?, ?, ?, ?, ?, ?)`
      )
      .run(
        input.candidateId,
        input.idempotencyKey,
        input.commandId,
        input.payloadDigest,
        JSON.stringify(input.receipt),
        Date.now()
      );
  }

  getIdempotentReceipt(
    candidateId: string,
    idempotencyKey: string
  ): { commandId: string; payloadDigest: string; receipt: MutationReceipt } | undefined {
    const row = this.db
      .prepare(
        "SELECT command_id, payload_digest, receipt FROM idempotency WHERE candidate_id = ? AND idempotency_key = ?"
      )
      .get(candidateId, idempotencyKey) as
      | { command_id: string; payload_digest: string; receipt: string }
      | undefined;
    if (!row) return undefined;
    return {
      commandId: row.command_id,
      payloadDigest: row.payload_digest,
      receipt: JSON.parse(row.receipt) as MutationReceipt
    };
  }

  countIdempotencyExecutions(candidateId: string, idempotencyKey: string): number {
    const row = this.db
      .prepare("SELECT COUNT(*) AS c FROM idempotency WHERE candidate_id = ? AND idempotency_key = ?")
      .get(candidateId, idempotencyKey) as { c: number };
    return row.c;
  }
}

function revisionFromRow(row: RevisionRow): CommittedRevision {
  return {
    revisionId: row.revision_id,
    sessionId: row.session_id,
    sequence: row.sequence,
    artifactRef: row.artifact_ref,
    contentHash: row.content_hash,
    origin: row.origin as RevisionOrigin,
    createdAt: row.created_at
  };
}

function candidateFromRow(row: CandidateRow): CandidateRevision {
  return {
    candidateId: row.candidate_id,
    sessionId: row.session_id,
    baseRevisionId: row.base_revision_id,
    baseHash: row.base_hash,
    artifactRef: row.artifact_ref,
    currentHash: row.current_hash ?? undefined,
    state: row.state as CandidateState,
    createdBy: row.created_by as "agent" | "human",
    verification: row.verification ? JSON.parse(row.verification) : undefined,
    failureReason: row.failure_reason ?? undefined,
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

export function leaseFromRow(row: LeaseRow): WriterLease {
  return {
    leaseId: row.lease_id,
    sessionId: row.session_id,
    owner: row.owner as WriterLease["owner"],
    backend: row.backend as WriterLease["backend"],
    baseRevisionId: row.base_revision_id,
    fencingToken: BigInt(row.fencing_token),
    sessionEpoch: row.epoch,
    acquiredAt: row.acquired_at
  };
}

function journalFromRow(row: JournalRow): CommitJournalRecord {
  return {
    commitId: row.commit_id,
    sessionId: row.session_id,
    candidateId: row.candidate_id,
    sourcePath: row.source_path,
    tempPath: row.temp_path,
    sourceHashBefore: row.source_hash_before,
    candidateHash: row.candidate_hash,
    phase: row.phase as CommitPhase,
    origin: row.origin as RevisionOrigin,
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}
