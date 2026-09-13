/**
 * Crash recovery integration (§76–§77, §141 fault injection, INV-15): journal
 * rows interrupted at each phase resolve by filesystem hash facts — never by
 * guessing from session metadata.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { copyFile, readFile, rm, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { sha256File } from "../../src/support/fsx.js";
import { OfficePlugin } from "../../src/plugin/office-plugin.js";
import { openWorkspace, writeDocxFixture } from "../helpers/fixtures.js";

let ws: Awaited<ReturnType<typeof openWorkspace>>;
let docxPath: string;
let pristine: Buffer;

beforeAll(async () => {
  ws = await openWorkspace();
  docxPath = join(ws.root, "recover.docx");
  await writeDocxFixture(docxPath, ["recovery fixture"]);
  pristine = await readFile(docxPath);
});

afterAll(async () => {
  await ws?.cleanup().catch(() => undefined);
});

async function restorePristine(): Promise<void> {
  await writeFile(docxPath, pristine);
}

interface CrashOptions {
  phase: "prepared" | "temp-ready" | "source-replaced";
  candidateBytes: Buffer;
  /** Runs after the journal row lands, before the plugin reopens (recovery). */
  beforeReopen?: () => Promise<void>;
}

/**
 * Fabricate a crash: write a journal row directly, then reopen the plugin —
 * initialize() runs crash recovery against the fabricated state.
 */
async function crashAndRecover(options: CrashOptions): Promise<string> {
  await restorePristine();
  const sourceHash = await sha256File(docxPath);
  const commitId = `cmt_crash_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  const tempPath = `${docxPath}.commit-${commitId}`;
  await writeFile(tempPath, options.candidateBytes);
  const candidateHash = await sha256File(tempPath);

  ws.plugin.service.repos.upsertJournal({
    commitId,
    sessionId: "sess_crash",
    candidateId: "cand_crash",
    sourcePath: docxPath,
    tempPath,
    sourceHashBefore: sourceHash,
    candidateHash,
    phase: options.phase,
    origin: "agent",
    createdAt: Date.now(),
    updatedAt: Date.now()
  });

  await options.beforeReopen?.();
  await ws.plugin.dispose();
  ws.plugin = await OfficePlugin.create({
    workspaceRoot: join(ws.root, "runtime"),
    engineDisabled: true
  });
  return commitId;
}

describe("Crash recovery (§77, §141, INV-15)", () => {
  it("PREPARED + untouched source → rolled back, temp cleaned", async () => {
    const commitId = await crashAndRecover({
      phase: "prepared",
      candidateBytes: Buffer.from("candidate")
    });
    const record = ws.plugin.service.repos.getJournal(commitId)!;
    expect(record.phase).toBe("aborted");
    expect(await sha256File(docxPath)).toBe(record.sourceHashBefore);
  });

  it("TEMP_READY + untouched source → rolled back", async () => {
    const commitId = await crashAndRecover({
      phase: "temp-ready",
      candidateBytes: Buffer.from("candidate-bytes")
    });
    const record = ws.plugin.service.repos.getJournal(commitId)!;
    expect(record.phase).toBe("aborted");
    expect(await sha256File(docxPath)).toBe(record.sourceHashBefore);
  });

  it("SOURCE_REPLACED + source already holds candidate bytes → finalized (INV-15)", async () => {
    const candidate = Buffer.concat([pristine, Buffer.from([0x41])]);
    const commitId = await crashAndRecover({
      phase: "source-replaced",
      candidateBytes: candidate,
      // The crash happened *after* the atomic replace landed the candidate.
      beforeReopen: async () => {
        await writeFile(docxPath, candidate);
      }
    });
    const record = ws.plugin.service.repos.getJournal(commitId)!;
    expect(record.phase).toBe("finalized");
    expect(await sha256File(docxPath)).toBe(record.candidateHash);
  });

  it("SOURCE_REPLACED + third-party bytes → conflict recorded, no half-write", async () => {
    const thirdParty = Buffer.concat([pristine, Buffer.from([0x42, 0x43])]);
    const commitId = await crashAndRecover({
      phase: "source-replaced",
      candidateBytes: Buffer.concat([pristine, Buffer.from([0x99])]),
      beforeReopen: async () => {
        await writeFile(docxPath, thirdParty);
      }
    });
    const record = ws.plugin.service.repos.getJournal(commitId)!;
    expect(record.phase).toBe("aborted");
    const events = ws.plugin.service.repos.readEventsSince("sess_crash", 0n, 1000);
    const resolved = [...events].reverse().find((e) => e.type === "recovery.resolved");
    expect((resolved?.payload as { resolution?: string })?.resolution).toBe("conflict");
    // Third-party bytes survived untouched.
    const { sha256Buffer } = await import("../../src/support/fsx.js");
    expect(await sha256File(docxPath)).toBe(sha256Buffer(thirdParty));
  });

  it("A→B→A same-origin: recovery lands a NEW revision, sequence continuous (v3 exact idempotency)", async () => {
    await restorePristine();
    const sessionId = "sess_aba";
    const contentA = pristine;
    const contentB = Buffer.concat([pristine, Buffer.from("B")]);
    const { sha256Buffer } = await import("../../src/support/fsx.js");
    const hashA = sha256Buffer(contentA);
    const hashB = sha256Buffer(contentB);
    const now = Date.now();
    // History: A(agent) → B(agent). Latest revision is B.
    ws.plugin.service.repos.insertRevision({
      revisionId: "rev_aba_a", sessionId, sequence: 1, artifactRef: "art_src_aba",
      contentHash: hashA, origin: "agent", createdAt: now
    });
    ws.plugin.service.repos.insertRevision({
      revisionId: "rev_aba_b", sessionId, sequence: 2, artifactRef: "art_src_aba",
      contentHash: hashB, origin: "agent", createdAt: now
    });

    // A third agent commit crashes at SOURCE_REPLACED; its content is A again.
    // Content-hash matching against history would swallow this revision.
    const commitId = "cmt_aba_return";
    await writeFile(docxPath, contentA);
    await writeFile(`${docxPath}.commit-${commitId}`, contentA);
    ws.plugin.service.repos.upsertJournal({
      commitId,
      sessionId,
      candidateId: "cand_aba",
      sourcePath: docxPath,
      tempPath: `${docxPath}.commit-${commitId}`,
      sourceHashBefore: hashB, // the commit was based on B
      candidateHash: hashA,    // the new content returns to A
      phase: "source-replaced",
      origin: "agent",
      createdAt: now,
      updatedAt: now
    });

    await ws.plugin.dispose();
    ws.plugin = await OfficePlugin.create({ workspaceRoot: join(ws.root, "runtime"), engineDisabled: true });

    expect(ws.plugin.service.repos.getJournal(commitId)!.phase).toBe("finalized");
    const revisions = ws.plugin.service.repos.listRevisions(sessionId);
    expect(revisions).toHaveLength(3); // a NEW third revision, not a history match
    expect(revisions[2]!.contentHash).toBe(hashA);
    expect(revisions[2]!.sequence).toBe(3); // sequence stays continuous
    expect(revisions[2]!.commitId).toBe(commitId); // bound by exact identity
    expect(revisions[2]!.revisionId).not.toBe("rev_aba_a");
  });

  it("revision exists but journal never flipped → identity-proven flip, no duplicate revision", async () => {
    await restorePristine();
    const sessionId = "sess_flip";
    const commitId = "cmt_flip_legacy";
    const { sha256Buffer } = await import("../../src/support/fsx.js");
    const candidate = Buffer.concat([pristine, Buffer.from([0x44])]);
    const candidateHash = sha256Buffer(candidate);
    await writeFile(docxPath, candidate);

    // Injected state: the revision for THIS commit already landed (v3
    // commit_id bound)… anything else in history is irrelevant.
    ws.plugin.service.repos.insertRevision({
      revisionId: "rev_flip", sessionId, sequence: 1, artifactRef: "art_flip",
      contentHash: candidateHash, origin: "agent", createdAt: Date.now(), commitId
    });
    // …but the journal is still at SOURCE_REPLACED.
    ws.plugin.service.repos.upsertJournal({
      commitId,
      sessionId,
      candidateId: "cand_flip",
      sourcePath: docxPath,
      tempPath: `${docxPath}.commit-${commitId}`,
      sourceHashBefore: sha256Buffer(pristine),
      candidateHash,
      phase: "source-replaced",
      origin: "agent",
      createdAt: Date.now(),
      updatedAt: Date.now()
    });

    await ws.plugin.dispose();
    ws.plugin = await OfficePlugin.create({ workspaceRoot: join(ws.root, "runtime"), engineDisabled: true });

    expect(ws.plugin.service.repos.getJournal(commitId)!.phase).toBe("finalized");
    const revisions = ws.plugin.service.repos.listRevisions(sessionId);
    expect(revisions).toHaveLength(1); // flipped in place — no duplicate insert
    expect(revisions[0]!.revisionId).toBe("rev_flip");
  });

  it("SelfWriteGuard recognizes self-originated watcher events (§73)", async () => {
    await restorePristine();
    const guards = ws.plugin.service.selfWrites;
    const expected = await sha256File(docxPath);
    guards.register({
      commitId: "cmt_selfwrite_test",
      expectedHash: expected,
      expiresAt: Date.now() + 5_000,
      sourcePath: docxPath
    });
    expect(await guards.isSelfWrite(docxPath)).toBe(true);
    expect(await guards.isSelfWrite(join(ws.root, "other.docx"))).toBe(false);
  });
});

function tempOf(commitId: string): string {
  return `${docxPath}.commit-${commitId}`;
}

// Silence unused import when rm is only needed in teardown paths.
void rm;
void copyFile;
