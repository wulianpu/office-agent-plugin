/**
 * Issue #9: cold-restart Session/Candidate/Watcher durable ownership.
 * Deterministic, engine-independent: fabricate persisted state, reopen the
 * plugin at the SAME workspace root, assert the reconciliation invariants.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { mkdir, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { OfficePlugin } from "../../src/plugin/office-plugin.js";
import { writeDocxFixture } from "../helpers/fixtures.js";

let root: string;

beforeAll(async () => {
  root = await mkdtemp(join(tmpdir(), "recover-own-"));
});

afterAll(async () => {
  await rm(root, { recursive: true, force: true }).catch(() => undefined);
});

async function waitFor(condition: () => boolean, deadlineMs = 10_000): Promise<void> {
  const deadline = Date.now() + deadlineMs;
  while (!condition()) {
    if (Date.now() > deadline) throw new Error("waitFor: condition not met in time");
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
}

describe("cold-restart ownership (issue #9)", () => {
  it("P1-high: a persisted READY candidate rebinds — Human/Agent gates agree; a second Agent task is blocked", async () => {
    const dir = join(root, "rebind");
    await mkdir(dir, { recursive: true });
    const docx = join(dir, "rebind.docx");
    await writeDocxFixture(docx, ["rebind fixture"]);
    const wsRoot = join(dir, "rt");
    const first = await OfficePlugin.create({ workspaceRoot: wsRoot, skipHostProbe: true });
    const ref = await first.registerArtifact(docx);
    const session = await first.openSession(ref);
    await first.service.sessions.ensureStrongIdentity(session.sessionId);
    const base = first.service.repos.latestRevision(session.sessionId)!;

    // Persist a crash-time READY candidate for this session.
    first.service.repos.upsertCandidate({
      candidateId: "cand_recovered_ready",
      sessionId: session.sessionId,
      baseRevisionId: base.revisionId,
      baseHash: base.contentHash,
      artifactRef: ref,
      currentHash: "ff".repeat(32),
      state: "ready",
      createdBy: "agent",
      verification: {
        contentHash: "ff".repeat(32),
        confidence: "verified",
        structural: { status: "pass", issues: [] }
      } as never,
      createdAt: Date.now(),
      updatedAt: Date.now()
    });
    await first.dispose();

    // Cold restart over the same DB.
    const second = await OfficePlugin.create({ workspaceRoot: wsRoot, skipHostProbe: true });
    try {
      const recovered = second.service.sessions.list().find(
        (s) => s.sessionId === session.sessionId
      );
      expect(recovered).toBeDefined();
      expect(recovered!.lifecycle).toBe("recovery-required");
      // The candidate was REBOUND to the session — one ownership fact.
      expect(recovered!.candidate?.candidateId).toBe("cand_recovered_ready");

      // Resolve to ready first (hash matches), then Agent begin must see
      // the durable candidate fact and refuse a second proposal.
      expect(await second.service.resolveRecoveredSession(session.sessionId)).toBe("ready");
      await expect(
        second.beginAgentTask(session.sessionId, { intent: "second", destructiveAllowed: false })
      ).rejects.toMatchObject({ code: "candidate-conflict" });

      // Accept of a foreign candidate id is rejected by ownership.
      await expect(
        second.service.acceptCandidate(session.sessionId, "cand_someone_else")
      ).rejects.toMatchObject({ code: "candidate-not-ready", message: expect.any(String) });
    } finally {
      await second.dispose().catch(() => undefined);
    }
  });

  it("P1-high: in-flight crash states never blind-resume — reconciled to failed with an explicit reason", async () => {
    const dir = join(root, "inflight");
    await mkdir(dir, { recursive: true });
    const docx = join(dir, "inflight.docx");
    await writeDocxFixture(docx, ["inflight fixture"]);
    const wsRoot = join(dir, "rt");
    const first = await OfficePlugin.create({ workspaceRoot: wsRoot, skipHostProbe: true });
    const ref = await first.registerArtifact(docx);
    const session = await first.openSession(ref);
    await first.service.sessions.ensureStrongIdentity(session.sessionId);
    const base = first.service.repos.latestRevision(session.sessionId)!;
    first.service.repos.upsertCandidate({
      candidateId: "cand_recovered_mutating",
      sessionId: session.sessionId,
      baseRevisionId: base.revisionId,
      baseHash: base.contentHash,
      artifactRef: ref,
      state: "mutating",
      createdBy: "agent",
      createdAt: Date.now(),
      updatedAt: Date.now()
    });
    await first.dispose();

    const second = await OfficePlugin.create({ workspaceRoot: wsRoot, skipHostProbe: true });
    try {
      const recovered = second.service.sessions.list().find(
        (s) => s.sessionId === session.sessionId
      )!;
      expect(recovered.candidate).toBeUndefined(); // not blindly rebound
      const row = second.service.repos.getCandidate("cand_recovered_mutating");
      expect(row?.state).toBe("failed");
      expect(row?.failureReason).toContain("recovered-after-crash");
      // The session is free to start a NEW task (no zombie blocking).
      expect(await second.service.resolveRecoveredSession(session.sessionId)).toBe("ready");
      const task = await second.beginAgentTask(session.sessionId, {
        intent: "fresh after crash",
        destructiveAllowed: false
      });
      await second.finalizeAgentTask(task);
      await second.rejectCandidate(session.sessionId);
      await second.closeSession(session.sessionId);
    } finally {
      await second.dispose().catch(() => undefined);
    }
  });

  it("P1-high: rehydrated sessions rejoin the SourceWatcher — external mutation flips a resolved session to conflict", async () => {
    const dir = join(root, "watcher");
    await mkdir(dir, { recursive: true });
    const docx = join(dir, "watcher.docx");
    await writeDocxFixture(docx, ["watcher fixture"]);
    const wsRoot = join(dir, "rt");
    const first = await OfficePlugin.create({ workspaceRoot: wsRoot, skipHostProbe: true });
    const ref = await first.registerArtifact(docx);
    const session = await first.openSession(ref);
    await first.service.sessions.ensureStrongIdentity(session.sessionId);
    await first.dispose();

    const { writeFile, appendFile, readFile } = await import("node:fs/promises");
    const second = await OfficePlugin.create({ workspaceRoot: wsRoot, skipHostProbe: true });
    try {
      const recovered = second.service.sessions.list().find(
        (s) => s.sessionId === session.sessionId
      )!;
      // Resolve back to ready first (hash matches the committed revision).
      const outcome = await second.service.resolveRecoveredSession(session.sessionId);
      expect(outcome).toBe("ready");

      // External save lands — the reattached watcher must flag the conflict.
      const pristine = await readFile(docx);
      await appendFile(docx, "x");
      await waitFor(
        () => second.service.getSession(session.sessionId)?.lifecycle === "conflict"
      );
      expect(second.service.getSession(session.sessionId)?.lifecycle).toBe("conflict");
      await writeFile(docx, pristine);
    } finally {
      await second.dispose().catch(() => undefined);
    }
  }, 20_000);

  it("P1: open-without-revision zombie rows are closed once at startup — no perpetual skip", async () => {
    const dir = join(root, "zombie");
    await mkdir(dir, { recursive: true });
    const docx = join(dir, "zombie.docx");
    await writeDocxFixture(docx, ["zombie fixture"]);
    const wsRoot = join(dir, "rt");
    const first = await OfficePlugin.create({ workspaceRoot: wsRoot, skipHostProbe: true });
    const ref = await first.registerArtifact(docx);
    const session = await first.openSession(ref);
    await first.service.sessions.ensureStrongIdentity(session.sessionId);
    // Simulate the crash window: insert a SECOND open session row with NO
    // revision materialized behind it.
    first.service.repos.upsertSession({
      sessionId: "sess_zombie_no_revision",
      documentId: "doc_zombie",
      artifactRef: ref,
      format: "docx",
      backend: "managed-file",
      lifecycle: "ready",
      epoch: 1,
      createdAt: Date.now()
    });
    await first.dispose();

    const second = await OfficePlugin.create({ workspaceRoot: wsRoot, skipHostProbe: true });
    try {
      const zombie = second.service.repos
        .listOpenSessions()
        .find((row) => row.session_id === "sess_zombie_no_revision");
      expect(zombie).toBeUndefined(); // closed at startup, not listed open
      const rehydrated = second.service.sessions.list().find(
        (s) => s.sessionId === session.sessionId
      );
      expect(rehydrated).toBeDefined(); // the real session still recovers
    } finally {
      await second.dispose().catch(() => undefined);
    }
  });
});
