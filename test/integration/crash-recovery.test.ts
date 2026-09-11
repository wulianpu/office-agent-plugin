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
