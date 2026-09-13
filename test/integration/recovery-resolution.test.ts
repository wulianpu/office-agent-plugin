/**
 * Recovery resolution + concurrent commit safety (§77, P11):
 * - rehydrated `recovery-required` sessions resolve to ready/conflict by hash facts
 * - two sessions on the same source file cannot interleave commits
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { appendFile } from "node:fs/promises";
import { join } from "node:path";
import { openWorkspace, writeDocxFixture } from "../helpers/fixtures.js";
import { createOfficeCliFixture } from "../helpers/officecli-fixture.js";

let ws: Awaited<ReturnType<typeof openWorkspace>>;
let fixture: Awaited<ReturnType<typeof createOfficeCliFixture>>;
let docxPath: string;

beforeAll(async () => {
  fixture = await createOfficeCliFixture();
  ws = await openWorkspace();
  if (fixture.available) {
    // Verification runs against the real engine — use a schema-valid fixture.
    docxPath = await fixture.docx(ws.root, "resolve.docx");
  } else {
    docxPath = join(ws.root, "resolve.docx");
    await writeDocxFixture(docxPath, ["resolution fixture"]);
  }
});

afterAll(async () => {
  await ws?.cleanup().catch(() => undefined);
});

describe("recovery resolution (§77)", () => {
  it("a rehydrated session with matching hash resolves to ready", async () => {
    const ref = await ws.plugin.registerArtifact(docxPath);
    const session = await ws.plugin.openSession(ref);
    const sessionId = session.sessionId;

    // Let the background strong identity bind first (rehydrated sessions in
    // practice always carry persisted, already-strong revisions).
    await ws.plugin.service.sessions.ensureStrongIdentity(sessionId);
    // Simulate the crash-recovery rehydration path: lifecycle → recovery-required.
    ws.plugin.service.sessions.updateSession(sessionId, (s) => {
      s.lifecycle = "recovery-required";
    });
    expect(ws.plugin.service.getSession(sessionId)?.lifecycle).toBe("recovery-required");

    const outcome = await ws.plugin.service.resolveRecoveredSession(sessionId);
    expect(outcome).toBe("ready");
    expect(ws.plugin.service.getSession(sessionId)?.lifecycle).toBe("ready");
    await ws.plugin.closeSession(sessionId);
  });

  it("a mutated source resolves to conflict, with a durable event", async () => {
    const ref = await ws.plugin.registerArtifact(docxPath);
    const session = await ws.plugin.openSession(ref);
    const sessionId = session.sessionId;

    // Bind strong identity on the ORIGINAL bytes first; external bytes then
    // appear before resolution (the divergence resolution must detect).
    await ws.plugin.service.sessions.ensureStrongIdentity(sessionId);
    const handle = await appendFile(docxPath, "x");
    void handle;
    ws.plugin.service.sessions.updateSession(sessionId, (s) => {
      s.lifecycle = "recovery-required";
    });

    const outcome = await ws.plugin.service.resolveRecoveredSession(sessionId);
    expect(outcome).toBe("conflict");
    expect(ws.plugin.service.getSession(sessionId)?.lifecycle).toBe("conflict");

    const events = ws.plugin.service.repos.readEventsSince(sessionId, 0n, 100);
    expect(events.some((e) => e.type === "session.conflict")).toBe(true);

    // Resolve is idempotent-guarded: not callable on a non-recovery session.
    await expect(ws.plugin.service.resolveRecoveredSession(sessionId)).rejects.toMatchObject({
      code: "recovery-required"
    });
    await ws.plugin.closeSession(sessionId);
  });
});

describe("concurrent commit safety (P11, INV-10/11)", () => {
  it("P0-4: a second session's writer on the same document is rejected at acquire", async () => {
    const ref = await ws.plugin.registerArtifact(docxPath);
    const sessionA = await ws.plugin.openSession(ref);
    const sessionB = await ws.plugin.openSession(ref);

    const taskA = await ws.plugin.beginAgentTask(sessionA.sessionId, {
      intent: "writer A",
      destructiveAllowed: false
    });
    // Document-scoped single writer: session B cannot acquire while A holds.
    await expect(
      ws.plugin.beginAgentTask(sessionB.sessionId, { intent: "writer B", destructiveAllowed: false })
    ).rejects.toMatchObject({ code: "lease-held" });
    // Human promotion is equally blocked.
    await expect(ws.plugin.beginEdit(sessionB.sessionId)).rejects.toMatchObject({ code: "lease-held" });

    // Release A; B may now acquire.
    await ws.plugin.finalizeAgentTask(taskA);
    await ws.plugin.rejectCandidate(sessionA.sessionId);
    const taskB = await ws.plugin.beginAgentTask(sessionB.sessionId, {
      intent: "writer B",
      destructiveAllowed: false
    });
    await ws.plugin.finalizeAgentTask(taskB);
    await ws.plugin.rejectCandidate(sessionB.sessionId);
    await ws.plugin.closeSession(sessionA.sessionId);
    await ws.plugin.closeSession(sessionB.sessionId);
  });

  // Engine-dependent: agent mutations + verification spawn OfficeCLI — a
  // runner without the engine must skip (spawn ENOENT), not fail. Runtime
  // skip: `it.skipIf` is evaluated at collection time, before beforeAll
  // assigns the fixture probe.
  it("serialized accepts on the same source file: exactly one wins", async (ctx) => {
    if (!fixture.available) ctx.skip();
    const ref = await ws.plugin.registerArtifact(docxPath);
    const sessionA = await ws.plugin.openSession(ref);
    const sessionB = await ws.plugin.openSession(ref);

    // Sequential tasks (document-scoped single writer forces ordering);
    // both candidates base on the same bytes with DIFFERENT mutations.
    const taskA = await ws.plugin.beginAgentTask(sessionA.sessionId, {
      intent: "concurrent A",
      destructiveAllowed: false
    });
    await ws.plugin.executeAgentMutation(taskA, {
      commandId: "cmd-cc-a",
      idempotencyKey: "cc-a",
      payload: [{ command: "set", path: "/body/paragraph[1]", props: { text: "Variant A" } }]
    });
    await ws.plugin.flushAgentCandidate(taskA);
    await ws.plugin.verifyAgentCandidate(taskA);
    await ws.plugin.finalizeAgentTask(taskA);

    const taskB = await ws.plugin.beginAgentTask(sessionB.sessionId, {
      intent: "concurrent B",
      destructiveAllowed: false
    });
    await ws.plugin.executeAgentMutation(taskB, {
      commandId: "cmd-cc-b",
      idempotencyKey: "cc-b",
      payload: [{ command: "set", path: "/body/paragraph[1]", props: { text: "Variant B" } }]
    });
    await ws.plugin.flushAgentCandidate(taskB);
    await ws.plugin.verifyAgentCandidate(taskB);
    await ws.plugin.finalizeAgentTask(taskB);
    // Snapshot hashes BEFORE accepting (the winner's candidate row is deleted on accept).
    const hashA = ws.plugin.service.candidates.require(taskA.candidateId).currentHash!;
    const hashB = ws.plugin.service.candidates.require(taskB.candidateId).currentHash!;

    // Fire both accepts concurrently — per-path serialization lets exactly one land.
    const results = await Promise.allSettled([
      ws.plugin.acceptCandidate(sessionA.sessionId, taskA.candidateId),
      ws.plugin.acceptCandidate(sessionB.sessionId, taskB.candidateId)
    ]);
    const fulfilled = results.filter((r) => r.status === "fulfilled");
    const rejected = results.filter((r) => r.status === "rejected");
    expect(fulfilled).toHaveLength(1);
    expect(rejected).toHaveLength(1);
    const reason = (rejected[0] as PromiseRejectedResult).reason as { code?: string; message?: string };
    expect(["source-mutated", "io-error", "candidate-not-ready"]).toContain(reason.code);

    // The surviving file is one of the two candidate byte sets — no interleave.
    const { sha256File } = await import("../../src/support/fsx.js");
    const finalHash = await sha256File(docxPath);
    expect([hashA, hashB]).toContain(finalHash);

    await ws.plugin.closeSession(sessionA.sessionId).catch(() => undefined);
    await ws.plugin.closeSession(sessionB.sessionId).catch(() => undefined);
  });
});

