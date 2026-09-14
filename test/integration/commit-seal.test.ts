/**
 * Issue #5 (P0): both destructive-replace seals are deterministic — the
 * candidate temp must hash to the verification-bound bytes BEFORE rename,
 * and the source gets a final INV-10 hash gate at the last moment. Neither
 * failure may touch the source, leave a revision, or register a self-write.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { mkdtemp, rm, writeFile, readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { RuntimeDatabase } from "../../src/runtime/persistence/database.js";
import { RuntimeRepositories } from "../../src/runtime/persistence/repositories.js";
import { DurableEventBus } from "../../src/runtime/sessions/event-bus.js";
import { RevisionLog } from "../../src/runtime/revisions/revision-log.js";
import { SelfWriteGuardRegistry } from "../../src/runtime/commit/self-write-guard.js";
import { AtomicFileCommitter } from "../../src/runtime/commit/atomic-file-committer.js";
import { sha256Buffer, sha256File } from "../../src/support/fsx.js";

let dir: string;
let db: RuntimeDatabase;
let repos: RuntimeRepositories;
let committer: AtomicFileCommitter;
let selfWrites: SelfWriteGuardRegistry;
let sourcePath: string;
let candidatePath: string;

const BASE = Buffer.from("base-content-v1");
const CANDIDATE = Buffer.from("candidate-content-v2");
const THIRD_PARTY = Buffer.from("external-save-during-commit-window");

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "commit-seal-"));
  db = new RuntimeDatabase(join(dir, "rt.db"));
  repos = new RuntimeRepositories(db);
  const events = new DurableEventBus(repos);
  const revisions = new RevisionLog(repos);
  selfWrites = new SelfWriteGuardRegistry();
  committer = new AtomicFileCommitter(repos, revisions, events, selfWrites);
  sourcePath = join(dir, "doc.docx");
  candidatePath = join(dir, "staging.docx");
  await writeFile(sourcePath, BASE);
  await writeFile(candidatePath, CANDIDATE);
});

afterAll(async () => {
  db?.close();
  await rm(dir, { recursive: true, force: true }).catch(() => undefined);
});

function request(overrides: { candidateHash?: string; expectedSourceHash?: string } = {}) {
  return {
    sessionId: "sess-seal",
    sessionEpoch: 1,
    candidateId: "cand-seal",
    sourcePath,
    candidatePath,
    sessionArtifactRef: "art-seal",
    expectedSourceHash: overrides.expectedSourceHash ?? sha256Buffer(BASE),
    candidateHash: overrides.candidateHash ?? sha256Buffer(CANDIDATE),
    origin: "agent" as const
  };
}

describe("pre-replace seals (issue #5, P0)", () => {
  it("source TOCTOU: an external save landing in the copy/fsync/journal window survives untouched", async () => {
    // Deterministic injection: at TEMP_READY (after the early hash check has
    // already passed) an external writer replaces the source bytes.
    committer.faultHook = async (phase) => {
      if (phase === "temp-ready") {
        await writeFile(sourcePath, THIRD_PARTY);
      }
    };
    try {
      await expect(committer.commit(request())).rejects.toMatchObject({ code: "source-mutated" });
    } finally {
      committer.faultHook = undefined;
    }
    // The external version is preserved verbatim — never overwritten.
    expect(await readFile(sourcePath)).toEqual(THIRD_PARTY);
    // Journal aborted; no revision; sibling temp cleaned; no self-write.
    const journal = repos.unresolvedJournal();
    expect(journal).toHaveLength(0);
    expect(repos.listJournal().every((j) => j.commitId.startsWith("cmt_") === false || j.phase !== "source-replaced")).toBe(true);
    expect(repos.listRevisions("sess-seal")).toHaveLength(0);
    expect(await readFile(sourcePath)).toEqual(THIRD_PARTY);
    expect(selfWrites.size).toBe(0);
  });

  it("candidate seal: staging bytes differing from the verification-bound hash fail BEFORE the replace", async () => {
    // The source is untouched by the previous test's external save: restore
    // the base bytes so only the candidate mismatch varies.
    await writeFile(sourcePath, BASE);
    const wrongHash = sha256Buffer(Buffer.from("some-other-bytes"));
    await expect(committer.commit(request({ candidateHash: wrongHash }))).rejects.toMatchObject({
      code: "candidate-hash-mismatch"
    });
    // Source stays at base; journal aborted; temp cleaned; nothing committed.
    expect(await sha256File(sourcePath)).toBe(sha256Buffer(BASE));
    expect(repos.listRevisions("sess-seal")).toHaveLength(0);
    expect(selfWrites.size).toBe(0);
    const all = repos.listJournal();
    expect(all.filter((j) => j.sessionId === "sess-seal").every((j) => j.phase === "aborted")).toBe(true);
  });

  it("source vanishing inside the window aborts with artifact-missing", async () => {
    committer.faultHook = async (phase) => {
      if (phase === "temp-ready") {
        await rm(sourcePath, { force: true });
      }
    };
    try {
      await expect(committer.commit(request())).rejects.toMatchObject({ code: "artifact-missing" });
    } finally {
      committer.faultHook = undefined;
      await writeFile(sourcePath, BASE); // restore for the happy-path test
    }
    expect(repos.listRevisions("sess-seal")).toHaveLength(0);
  });

  it("happy path: both seals pass, the commit finalizes with identity binding", async () => {
    const result = await committer.commit(request());
    expect(result.journalOutcome).toBe("finalized");
    expect(result.newRevision.contentHash).toBe(sha256Buffer(CANDIDATE));
    expect(await sha256File(sourcePath)).toBe(sha256Buffer(CANDIDATE));
    expect(selfWrites.size).toBe(1); // registered only on success
    expect(repos.listRevisions("sess-seal")).toHaveLength(1);
  });
});
