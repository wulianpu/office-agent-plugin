/**
 * Round 8 (issue #2): schema migration crash-atomicity + compatibility-lock
 * invariants. Every migration version commits its DDL together with the
 * schemaVersion bump; a legacy partial state (column added, meta stale) must
 * recover on reopen instead of bricking on duplicate column/index.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { DB_SCHEMA_VERSION, RuntimeDatabase } from "../../src/runtime/persistence/database.js";
import { RuntimeRepositories } from "../../src/runtime/persistence/repositories.js";
import { COMPATIBILITY_LOCK } from "../../src/plugin/office-plugin.js";

let dir: string;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "dbmigrate-test-"));
});

afterAll(async () => {
  await rm(dir, { recursive: true, force: true }).catch(() => undefined);
});

/**
 * Fabricate a physical v2 DB: build the current schema, then downgrade
 * (drop the v3 index + column, rewind meta) and seed a legacy revision row.
 * `partialV3Column` additionally simulates a legacy pre-atomic crash: the
 * commit_id column exists but the index/meta never landed.
 */
function fabricateV2(dbPath: string, options: { partialV3Column?: boolean } = {}): void {
  const db = new RuntimeDatabase(dbPath);
  db.db.exec("DROP INDEX IF EXISTS idx_revisions_commit_id");
  db.db.exec("ALTER TABLE revisions DROP COLUMN commit_id");
  db.db.prepare("UPDATE meta SET value = '2' WHERE key = 'schemaVersion'").run();
  db.db
    .prepare(
      "INSERT INTO revisions (revision_id, session_id, sequence, artifact_ref, content_hash, origin, created_at) VALUES ('rev_legacy', 'sess_v2', 1, 'art_v2', 'aa', 'agent', 123)"
    )
    .run();
  if (options.partialV3Column) {
    db.db.exec("ALTER TABLE revisions ADD COLUMN commit_id TEXT");
  }
  db.close();
}

function schemaVersionOf(db: RuntimeDatabase): number {
  const row = db.db.prepare("SELECT value FROM meta WHERE key = 'schemaVersion'").get() as {
    value: string;
  };
  return Number(row.value);
}

describe("schema migration atomicity (round 8, issue #2)", () => {
  it("compatibility contract never drifts from the persistence schema", () => {
    expect(COMPATIBILITY_LOCK.dbSchema).toBe(DB_SCHEMA_VERSION);
  });

  it("fresh DB migrates 0 → current atomically; reopen is a no-op", () => {
    const db = new RuntimeDatabase(join(dir, "fresh.db"));
    expect(schemaVersionOf(db)).toBe(DB_SCHEMA_VERSION);
    db.close();
    const again = new RuntimeDatabase(join(dir, "fresh.db"));
    expect(schemaVersionOf(again)).toBe(DB_SCHEMA_VERSION);
    again.close();
  });

  it("v2 → v3: legacy revisions keep NULL commit_id; identity lookups work", () => {
    const dbPath = join(dir, "v2clean.db");
    fabricateV2(dbPath);
    const db = new RuntimeDatabase(dbPath);
    expect(schemaVersionOf(db)).toBe(DB_SCHEMA_VERSION);
    const repos = new RuntimeRepositories(db);
    const revisions = repos.listRevisions("sess_v2");
    expect(revisions).toHaveLength(1);
    expect(revisions[0]!.commitId).toBeUndefined(); // legacy row: NULL, not guessed
    expect(repos.findRevisionByCommitId("cmt_x")).toBeUndefined();
    db.db
      .prepare(
        "INSERT INTO revisions (revision_id, session_id, sequence, artifact_ref, content_hash, origin, created_at, commit_id) VALUES ('rev_new', 'sess_v2', 2, 'art_v2', 'bb', 'agent', 124, 'cmt_x')"
      )
      .run();
    expect(repos.findRevisionByCommitId("cmt_x")?.revisionId).toBe("rev_new");
    db.close();
  });

  it("interrupted migration (column exists, meta stale at 2): reopen completes, never bricks", () => {
    const dbPath = join(dir, "v2partial.db");
    fabricateV2(dbPath, { partialV3Column: true });
    // Pre-fix behavior: ALTER TABLE ... ADD COLUMN commit_id throws
    // "duplicate column name" and the runtime cannot start at all.
    const db = new RuntimeDatabase(dbPath);
    expect(schemaVersionOf(db)).toBe(DB_SCHEMA_VERSION);
    const index = db.db
      .prepare("SELECT name FROM sqlite_master WHERE type = 'index' AND name = 'idx_revisions_commit_id'")
      .get();
    expect(index).toBeTruthy();
    const repos = new RuntimeRepositories(db);
    expect(repos.listRevisions("sess_v2")).toHaveLength(1);
    expect(repos.listRevisions("sess_v2")[0]!.revisionId).toBe("rev_legacy");
    db.close();
  });
});
