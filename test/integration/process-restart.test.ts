const PROJECT_ROOT = process.cwd();
/**
 * Process-restart integration test (§77, §141): the REAL cross-process
 * recovery chain — kill → new process → hydrate → recover → resolve refs.
 *
 * Uses child processes to prove the recovery works with cold memory (no
 * shared in-process state between the "crash" and the "restart").
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { mkdtemp, rm, mkdir, readFile, writeFile, copyFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { createHash } from "node:crypto";

const run = promisify(execFile);
const sha = (data: Buffer) => createHash("sha256").update(data).digest("hex");

let workspace: string;

beforeAll(async () => {
  workspace = await mkdtemp(join(tmpdir(), "restart-test-"));
});

afterAll(async () => {
  await rm(workspace, { recursive: true, force: true }).catch(() => undefined);
});

/**
 * Phase-1 script: runs inside a child process. Creates a workspace, commits
 * a revision, then writes the commit phase to simulate different crash
 * states. Exits with the workspace path + artifactRef printed on stdout.
 */
const PHASE1_SCRIPT = `
const { OfficePlugin } = await import((await import("node:url")).pathToFileURL(process.cwd() + "/dist/plugin/office-plugin.js").href);
const { writeFile, mkdir } = await import("node:fs/promises");
const path = await import("node:path");

const ws = process.argv[2];
const mode = process.argv[3]; // "committed" | "crash-prepared" | "crash-source-replaced"
const plugin = await OfficePlugin.create({ workspaceRoot: path.join(ws, "rt"), skipHostProbe: true });

// Create a real docx source file
const { execFile } = await import("node:child_process");
const { promisify } = await import("node:util");
const exec = promisify(execFile);
const srcPath = path.join(ws, "source.docx");
await exec("officecli", ["create", srcPath, "--json"]).catch(() => undefined);
await exec("officecli", ["batch", srcPath, "--commands", JSON.stringify([
  { command: "add", parent: "/body", type: "paragraph", props: { text: "Restart test content" } }
]), "--json"]);
await exec("officecli", ["close", srcPath, "--json"]);

const ref = await plugin.registerArtifact(srcPath);
const session = await plugin.openSession(ref);
await plugin.service.sessions.ensureStrongIdentity(session.sessionId);

if (mode === "committed") {
  const task = await plugin.beginAgentTask(session.sessionId, { intent: "commit", destructiveAllowed: false });
  await plugin.executeAgentMutation(task, {
    commandId: "cmd-r1",
    idempotencyKey: "restart-1",
    payload: [{ command: "set", path: "/body/paragraph[1]", props: { text: "Committed Version" } }]
  });
  await plugin.flushAgentCandidate(task);
  await plugin.verifyAgentCandidate(task);
  await plugin.finalizeAgentTask(task);
  const result = await plugin.acceptCandidate(session.sessionId);
  console.log(JSON.stringify({ ref, revisionId: result.revisionId, mode }));
} else if (mode === "crash-prepared") {
  // Fabricate a journal row at PREPARED with an untouched source
  const { sha256File } = await import((await import("node:url")).pathToFileURL(process.cwd() + "/dist/support/fsx.js").href);
  const sourceHash = await sha256File(srcPath);
  const repos = plugin.service.repos;
  repos.upsertJournal({
    commitId: "cmt_crash_prep",
    sessionId: session.sessionId,
    candidateId: "cand_crash_prep",
    sourcePath: srcPath,
    tempPath: srcPath + ".commit-cmt_crash_prep",
    sourceHashBefore: sourceHash,
    candidateHash: "ff".repeat(32),
    phase: "prepared",
    origin: "agent",
    createdAt: Date.now(),
    updatedAt: Date.now()
  });
  console.log(JSON.stringify({ ref, revisionId: session.committedRevision.revisionId, mode }));
} else if (mode === "crash-source-replaced") {
  // Replace source with candidate bytes, leave journal at SOURCE_REPLACED
  const { sha256File } = await import((await import("node:url")).pathToFileURL(process.cwd() + "/dist/support/fsx.js").href);
  const sourceHash = await sha256File(srcPath);
  const candidateBytes = Buffer.from([0x50, 0x4b, 0x03, 0x04, ...new Uint8Array(200)]);
  const tempPath = srcPath + ".commit-cmt_crash_sr";
  await writeFile(tempPath, candidateBytes);
  const { rename } = await import("node:fs/promises");
  await rename(tempPath, srcPath);
  const candidateHash = await sha256File(srcPath);
  const repos = plugin.service.repos;
  repos.upsertJournal({
    commitId: "cmt_crash_sr",
    sessionId: session.sessionId,
    candidateId: "cand_crash_sr",
    sourcePath: srcPath,
    tempPath: tempPath,
    sourceHashBefore: sourceHash,
    candidateHash: candidateHash,
    phase: "source-replaced",
    origin: "agent",
    createdAt: Date.now(),
    updatedAt: Date.now()
  });
  console.log(JSON.stringify({ ref, revisionId: null, mode, candidateHash }));
}

await plugin.dispose();
process.exit(0);
`;

const PHASE2_SCRIPT = `
const { OfficePlugin } = await import((await import("node:url")).pathToFileURL(process.cwd() + "/dist/plugin/office-plugin.js").href);
const path = await import("node:path");

const ws = process.argv[2];
const plugin = await OfficePlugin.create({ workspaceRoot: path.join(ws, "rt"), skipHostProbe: true });

const recovery = await plugin.recover();
const sessions = plugin.service.sessions.list();
const artifacts = plugin.service.repos.loadArtifacts();

// Verify ArtifactRef resolution (the P0-1 test: hydrated store)
const refResults = [];
for (const a of artifacts.filter(x => x.kind === "source")) {
  try {
    const p = plugin.service.store.resolvePath(a.ref);
    refResults.push({ ref: a.ref, resolved: true, path: p });
  } catch (e) {
    refResults.push({ ref: a.ref, resolved: false });
  }
}

const journal = plugin.service.repos.listJournal();
const revisions = plugin.service.repos.listJournal().map(j => ({
  commitId: j.commitId,
  phase: j.phase
}));

console.log(JSON.stringify({
  recovery,
  refResults,
  journal: revisions,
  artifactCount: artifacts.length
}));

await plugin.dispose();
process.exit(0);
`;

async function runPhase(script: string, ...args: string[]): Promise<Record<string, unknown>> {
  const scriptPath = join(workspace, `phase-${Date.now()}-${Math.random().toString(36).slice(2)}.mjs`);
  await writeFile(scriptPath, script, "utf8");
  try {
    const { stdout } = await run(process.execPath, [scriptPath, ...args], {
      cwd: process.cwd(),
      windowsHide: true,
      timeout: 120_000
    });
    const jsonLine = stdout.trim().split("\n").find((l) => l.startsWith("{"));
    return JSON.parse(jsonLine ?? "{}");
  } finally {
    await rm(scriptPath, { force: true }).catch(() => undefined);
  }
}

/** CI without the OfficeCLI engine: the child processes need it for fixtures. */
const engineAvailable = await import("../../src/agent/officecli/officecli-adapter.js")
  .then(async (m) => {
    const probe = new m.OfficeCliAdapter();
    return probe.version_().then(() => true).catch(() => false);
  })
  .catch(() => false);

describe.skipIf(!engineAvailable)("process-restart recovery (§77)", () => {
  it("committed state: restart resolves all refs and revisions intact", async () => {
    const dir = join(workspace, "committed");
    await mkdir(dir, { recursive: true });
    const phase1 = await runPhase(PHASE1_SCRIPT, dir, "committed");

    expect(phase1.ref).toBeTruthy();
    expect(phase1.revisionId).toBeTruthy();

    const phase2 = await runPhase(PHASE2_SCRIPT, dir);

    // ArtifactRef resolves (P0-1 hydration test)
    const sourceRef = (phase2.refResults as Array<{ ref: string; resolved: boolean }>).find(
      (r: { ref: string; resolved: boolean }) => (r as { ref: string }).ref === phase1.ref
    );
    expect(sourceRef).toBeDefined();
    expect((sourceRef as { resolved: boolean }).resolved).toBe(true);

    // No unresolved journal entries
    const unresolved = (phase2.journal as Array<{ phase: string }>).filter(
      (j) => j.phase !== "finalized" && j.phase !== "aborted"
    );
    expect(unresolved).toHaveLength(0);
  }, 300_000);

  it("PREPARED crash: restart recovers → rolled back, ref resolvable", async () => {
    const dir = join(workspace, "prepared");
    await mkdir(dir, { recursive: true });
    await runPhase(PHASE1_SCRIPT, dir, "crash-prepared");

    const phase2 = await runPhase(PHASE2_SCRIPT, dir);

    const prep = (phase2.journal as Array<{ commitId: string; phase: string }>).find(
      (j) => j.commitId === "cmt_crash_prep"
    );
    expect(prep?.phase).toBe("aborted"); // source untouched → rollback

    // Source artifact still resolvable
    const resolved = (phase2.refResults as Array<{ resolved: boolean }>).filter((r) => r.resolved);
    expect(resolved.length).toBeGreaterThan(0);
  }, 300_000);

  it("SOURCE_REPLACED crash: restart forward-recovers revision", async () => {
    const dir = join(workspace, "source-replaced");
    await mkdir(dir, { recursive: true });
    const phase1 = await runPhase(PHASE1_SCRIPT, dir, "crash-source-replaced");

    const phase2 = await runPhase(PHASE2_SCRIPT, dir);

    const sr = (phase2.journal as Array<{ commitId: string; phase: string }>).find(
      (j) => j.commitId === "cmt_crash_sr"
    );
    expect(sr?.phase).toBe("finalized"); // source holds candidate → forward recover

    // Refs resolve
    const resolved = (phase2.refResults as Array<{ resolved: boolean }>).filter((r) => r.resolved);
    expect(resolved.length).toBeGreaterThan(0);
  }, 300_000);
});
