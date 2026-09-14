/**
 * Issue #5 reopen: the pre-replace source seal must gate EVERY Windows
 * rename retry attempt (the 120..1200ms waits + lock releaser re-open a
 * multi-second TOCTOU), and commit durability claims must align
 * (synchronous=FULL + temp-ready phase-regression forward recovery).
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { mkdtemp, rm, writeFile, readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { DatabaseSync } from "node:sqlite";
import { rename } from "node:fs/promises";
import { RuntimeDatabase } from "../../src/runtime/persistence/database.js";
import { RuntimeRepositories } from "../../src/runtime/persistence/repositories.js";
import { DurableEventBus } from "../../src/runtime/sessions/event-bus.js";
import { RevisionLog } from "../../src/runtime/revisions/revision-log.js";
import { SelfWriteGuardRegistry } from "../../src/runtime/commit/self-write-guard.js";
import { AtomicFileCommitter } from "../../src/runtime/commit/atomic-file-committer.js";
import { sha256Buffer } from "../../src/support/fsx.js";

let dir: string;
let db: RuntimeDatabase;
let repos: RuntimeRepositories;
let committer: SeamCommitter;
let selfWrites: SelfWriteGuardRegistry;
let sourcePath: string;
let candidatePath: string;

const BASE = Buffer.from("retry-base-v1");
const CANDIDATE = Buffer.from("retry-candidate-v2");
const EXTERNAL = Buffer.from("external-save-during-retry-wait");

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "retry-seal-"));
  sourcePath = join(dir, "doc.docx");
  candidatePath = join(dir, "staging.docx");
  await writeFile(sourcePath, BASE);
  await writeFile(candidatePath, CANDIDATE);
});

afterAll(async () => {
  db?.close();
  await rm(dir, { recursive: true, force: true }).catch(() => undefined);
});

/** Exposes the protected rename seam for deterministic fault injection. */
class SeamCommitter extends AtomicFileCommitter {
  setRename(fn: (temp: string, source: string) => Promise<void>): void {
    this.renameAttempt = fn;
  }
  static realRename(): (temp: string, source: string) => Promise<void> {
    return (temp, source) => rename(temp, source);
  }
}

function freshCommitter(): SeamCommitter {
  db = new RuntimeDatabase(join(dir, `rt-${Date.now()}-${Math.random().toString(36).slice(2)}.db`));
  repos = new RuntimeRepositories(db);
  const events = new DurableEventBus(repos);
  const revisions = new RevisionLog(repos);
  selfWrites = new SelfWriteGuardRegistry();
  return new SeamCommitter(repos, revisions, events, selfWrites);
}

function request() {
  return {
    sessionId: "sess-retry",
    sessionEpoch: 1,
    candidateId: "cand-retry",
    sourcePath,
    candidatePath,
    sessionArtifactRef: "art-retry",
    expectedSourceHash: sha256Buffer(BASE),
    candidateHash: sha256Buffer(CANDIDATE),
    origin: "agent" as const
  };
}

describe("per-attempt replace seal (issue #5 reopen, P0-1)", () => {
  it("an external save landing during a retry wait is preserved — rename attempt 2 never executes", async () => {
    committer = freshCommitter();
    let attempts = 0;
    committer.setRename(async (temp: string, source: string) => {
      void temp;
      attempts++;
      if (attempts === 1) {
        // Attempt 1 hits a Windows-style lock failure; the external writer
        // saves DURING the retry wait that follows.
        await writeFile(source, EXTERNAL);
        throw Object.assign(new Error("file locked"), { code: "EPERM" });
      }
      throw new Error("attempt 2 must never run: the seal must stop the retries");
    });
    await expect(committer.commit(request())).rejects.toMatchObject({ code: "source-mutated" });
    expect(attempts).toBe(1); // retries stopped at the per-attempt seal
    expect(await readFile(sourcePath)).toEqual(EXTERNAL); // external bytes preserved verbatim
    const all = repos.listJournal().filter((j) => j.sessionId === "sess-retry");
    expect(all.every((j) => j.phase === "aborted")).toBe(true);
    expect(repos.listRevisions("sess-retry")).toHaveLength(0);
    expect(selfWrites.size).toBe(0);
  });

  it("lock-resilient happy path: a transient EPERM without external changes still succeeds on retry", async () => {
    await writeFile(sourcePath, BASE); // restore base
    committer = freshCommitter();
    const realRename = SeamCommitter.realRename();
    let attempts = 0;
    committer.setRename(async (temp: string, source: string) => {
      attempts++;
      if (attempts === 1) throw Object.assign(new Error("locked"), { code: "EPERM" });
      return realRename(temp, source);
    });
    const result = await committer.commit(request());
    expect(attempts).toBe(2); // retried and succeeded
    expect(result.journalOutcome).toBe("finalized");
    expect(selfWrites.size).toBe(1);
  });
});

describe("durability alignment (issue #5 reopen, P0-2)", () => {
  it("commit journal runs at synchronous=FULL — same power-loss durability class as the source fsync", () => {
    const raw = new DatabaseSync(join(dir, "pragma.db"));
    raw.exec("PRAGMA journal_mode = WAL");
    raw.exec("PRAGMA synchronous = FULL");
    const mode = raw.prepare("PRAGMA synchronous").get() as { synchronous: number };
    expect(mode.synchronous).toBe(2); // 2 = FULL
    raw.close();
    // And the runtime database itself opens with FULL.
    const runtime = new RuntimeDatabase(join(dir, "pragma-rt.db"));
    const value = (runtime.db.prepare("PRAGMA synchronous").get() as { synchronous: number }).synchronous;
    expect(value).toBe(2);
    runtime.close();
  });

  it("phase regression: journal=temp-ready + source already holds candidate bytes forward-finalizes with exact commit_id", async () => {
    await writeFile(sourcePath, BASE);
    committer = freshCommitter();
    // Synthesize the regression state: durable phase lagged the filesystem
    // (pre-FULL durability split / injected legacy state).
    const commitId = "cmt_regression_temp_ready";
    await writeFile(sourcePath, CANDIDATE); // filesystem already replaced
    repos.upsertJournal({
      commitId,
      sessionId: "sess-retry",
      candidateId: "cand-retry",
      sourcePath,
      tempPath: `${sourcePath}.commit-${commitId}`,
      sourceHashBefore: sha256Buffer(BASE),
      candidateHash: sha256Buffer(CANDIDATE),
      phase: "temp-ready",
      origin: "agent",
      createdAt: Date.now(),
      updatedAt: Date.now()
    });
    const { RecoveryService } = await import("../../src/runtime/recovery/recovery-service.js");
    const events = new DurableEventBus(repos);
    const revisions = new RevisionLog(repos);
    const store = {
      tryResolveRefByPath: () => "art-retry",
      register: async () => "art-retry"
    } as never;
    const recovery = new RecoveryService(repos, revisions, events, committer, store);
    const outcomes = await recovery.recoverAll();
    const outcome = outcomes.find((o) => o.commitId === commitId);
    expect(outcome?.resolution).toBe("finalized"); // NOT conflict/aborted
    expect(repos.getJournal(commitId)?.phase).toBe("finalized");
    const landed = repos.findRevisionByCommitId(commitId);
    expect(landed).toBeTruthy(); // exact commit_id binding preserved
    expect(landed?.contentHash).toBe(sha256Buffer(CANDIDATE));
  });
});
