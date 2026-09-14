/**
 * Persistence (§115, §119): SQLite via node:sqlite, WAL mode, owned solely by
 * OfficeRuntimeService. Transactions are short and synchronous — no read
 * transaction is ever held across editor lifetimes. bigint values (fencing
 * tokens, event sequences) are stored as TEXT because node:sqlite rejects
 * integers beyond MAX_SAFE_INTEGER.
 */

import { DatabaseSync } from "node:sqlite";

export const DB_SCHEMA_VERSION = 3;

export class RuntimeDatabase {
  readonly db: DatabaseSync;

  constructor(readonly dbPath: string) {
    this.db = new DatabaseSync(dbPath);
    this.db.exec("PRAGMA journal_mode = WAL");
    // P0-2 (#5 reopen): FULL — commit-journal phases must carry the same
    // power-loss durability claim as the committer's post-rename fsync of
    // the source file. NORMAL only fsyncs at checkpoints: a machine crash
    // could leave the filesystem durable with the candidate bytes while the
    // last SOURCE_REPLACED WAL transaction was still lost — exactly the
    // split Recovery must never have to guess about.
    this.db.exec("PRAGMA synchronous = FULL");
    this.db.exec("PRAGMA busy_timeout = 5000");
    this.db.exec("PRAGMA foreign_keys = ON");
    this.migrate();
  }

  private migrate(): void {
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS meta (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL
      );
    `);
    const row = this.db
      .prepare("SELECT value FROM meta WHERE key = 'schemaVersion'")
      .get() as { value: string } | undefined;
    const current = row ? Number(row.value) : 0;
    if (current > DB_SCHEMA_VERSION) {
      throw new Error(
        `session DB schema ${current} is newer than runtime supports (${DB_SCHEMA_VERSION})`
      );
    }
    // Crash-atomic per version (round 8, P1-high): each migration's DDL and
    // the schemaVersion bump share ONE transaction. A crash mid-migration
    // rolls the DDL back, so a restart retries exactly this version — never
    // a half-applied schema with a stale meta. The old code flipped meta only
    // after ALL versions, leaving a duplicate-column brick window.
    for (let v = current + 1; v <= DB_SCHEMA_VERSION; v++) {
      this.withTransaction(() => {
        this.applyMigration(v);
        this.db
          .prepare("INSERT OR REPLACE INTO meta (key, value) VALUES ('schemaVersion', ?)")
          .run(String(v));
      });
    }
  }

  private applyMigration(version: number): void {
    if (version === 1) {
      this.db.exec(`
        CREATE TABLE artifacts (
          ref TEXT PRIMARY KEY,
          path TEXT NOT NULL UNIQUE,
          kind TEXT NOT NULL,
          format TEXT NOT NULL,
          created_at INTEGER NOT NULL
        );
        CREATE TABLE sessions (
          session_id TEXT PRIMARY KEY,
          document_id TEXT NOT NULL,
          artifact_ref TEXT NOT NULL,
          format TEXT NOT NULL,
          backend TEXT NOT NULL,
          lifecycle TEXT NOT NULL,
          epoch INTEGER NOT NULL,
          created_at INTEGER NOT NULL,
          closed_at INTEGER
        );
        CREATE TABLE revisions (
          revision_id TEXT PRIMARY KEY,
          session_id TEXT NOT NULL,
          sequence INTEGER NOT NULL,
          artifact_ref TEXT NOT NULL,
          content_hash TEXT NOT NULL,
          origin TEXT NOT NULL,
          created_at INTEGER NOT NULL
        );
        CREATE INDEX idx_revisions_session ON revisions(session_id, sequence);
        CREATE TABLE candidates (
          candidate_id TEXT PRIMARY KEY,
          session_id TEXT NOT NULL,
          base_revision_id TEXT NOT NULL,
          base_hash TEXT NOT NULL,
          artifact_ref TEXT NOT NULL,
          current_hash TEXT,
          state TEXT NOT NULL,
          created_by TEXT NOT NULL,
          verification TEXT,
          failure_reason TEXT,
          created_at INTEGER NOT NULL,
          updated_at INTEGER NOT NULL
        );
        CREATE TABLE leases (
          lease_id TEXT PRIMARY KEY,
          session_id TEXT NOT NULL,
          owner TEXT NOT NULL,
          backend TEXT NOT NULL,
          base_revision_id TEXT NOT NULL,
          fencing_token TEXT NOT NULL,
          epoch INTEGER NOT NULL,
          acquired_at INTEGER NOT NULL,
          released_at INTEGER
        );
        CREATE TABLE events (
          event_id TEXT PRIMARY KEY,
          session_id TEXT NOT NULL,
          sequence TEXT NOT NULL,
          epoch INTEGER NOT NULL,
          type TEXT NOT NULL,
          payload TEXT NOT NULL,
          created_at INTEGER NOT NULL
        );
        CREATE INDEX idx_events_session ON events(session_id, sequence);
        CREATE TABLE commit_journal (
          commit_id TEXT PRIMARY KEY,
          session_id TEXT NOT NULL,
          candidate_id TEXT NOT NULL,
          source_path TEXT NOT NULL,
          temp_path TEXT NOT NULL,
          source_hash_before TEXT NOT NULL,
          candidate_hash TEXT NOT NULL,
          phase TEXT NOT NULL,
          origin TEXT NOT NULL,
          created_at INTEGER NOT NULL,
          updated_at INTEGER NOT NULL
        );
        CREATE TABLE idempotency (
          session_id TEXT NOT NULL,
          idempotency_key TEXT NOT NULL,
          command_id TEXT NOT NULL,
          receipt TEXT NOT NULL,
          created_at INTEGER NOT NULL,
          PRIMARY KEY (session_id, idempotency_key)
        );
        CREATE TABLE scans (
          content_hash TEXT PRIMARY KEY,
          size TEXT NOT NULL,
          manifest TEXT NOT NULL,
          corrupt_entries TEXT NOT NULL,
          scanned_at INTEGER NOT NULL
        );
      `);
    }
    if (version === 2) {
      // v2 (§61): idempotency moves from (session_id, key) to (candidate_id,
      // key) with a payload digest — retries are per-task; cross-task key
      // reuse with a different payload raises IDEMPOTENCY_CONFLICT.
      this.db.exec(`
        DROP TABLE IF EXISTS idempotency;
        CREATE TABLE idempotency (
          candidate_id TEXT NOT NULL,
          idempotency_key TEXT NOT NULL,
          command_id TEXT NOT NULL,
          payload_digest TEXT NOT NULL,
          receipt TEXT NOT NULL,
          created_at INTEGER NOT NULL,
          PRIMARY KEY (candidate_id, idempotency_key)
        );
      `);
    }
    if (version === 3) {
      // v3 (§76 exact-commit idempotency): revisions carry the commit_id of
      // the journal row that produced them. UNIQUE (NULLs allowed for legacy
      // rows and journal-less external revisions) — recovery decides "has
      // this commit landed?" by exact identity, never by content hash.
      // Tolerant to a legacy partial state (column added by a pre-atomic
      // migration that crashed before flipping meta): re-running must
      // complete, never brick on duplicate column/index.
      const columns = this.db.prepare("PRAGMA table_info(revisions)").all() as Array<{
        name: string;
      }>;
      if (!columns.some((c) => c.name === "commit_id")) {
        this.db.exec(`ALTER TABLE revisions ADD COLUMN commit_id TEXT;`);
      }
      const index = this.db
        .prepare(
          "SELECT name FROM sqlite_master WHERE type = 'index' AND name = 'idx_revisions_commit_id'"
        )
        .get();
      if (!index) {
        this.db.exec(`CREATE UNIQUE INDEX idx_revisions_commit_id ON revisions(commit_id);`);
      }
    }
  }

  /** Short synchronous transaction (§119): result copied out before COMMIT. */
  withTransaction<T>(fn: () => T): T {
    this.db.exec("BEGIN IMMEDIATE");
    try {
      const result = fn();
      this.db.exec("COMMIT");
      return result;
    } catch (error) {
      this.db.exec("ROLLBACK");
      throw error;
    }
  }

  setMeta(key: string, value: string): void {
    this.db
      .prepare("INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)")
      .run(key, value);
  }

  getMeta(key: string): string | undefined {
    const row = this.db.prepare("SELECT value FROM meta WHERE key = ?").get(key) as
      | { value: string }
      | undefined;
    return row?.value;
  }

  close(): void {
    this.db.close();
  }
}
