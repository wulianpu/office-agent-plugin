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

    // External bytes appear before resolution.
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
  it("concurrent accepts on the same source file: exactly one wins", async () => {
    const ref = await ws.plugin.registerArtifact(docxPath);
    // Two independent sessions on the same artifact.
    const sessionA = await ws.plugin.openSession(ref);
    const sessionB = await ws.plugin.openSession(ref);

    const taskA = await ws.plugin.beginAgentTask(sessionA.sessionId, {
      intent: "concurrent A",
      destructiveAllowed: false
    });
    const taskB = await ws.plugin.beginAgentTask(sessionB.sessionId, {
      intent: "concurrent B",
      destructiveAllowed: false
    });
    // Both candidates base on the same bytes but carry DIFFERENT mutations —
    // only one can survive a replace.
    await ws.plugin.executeAgentMutation(taskA, {
      commandId: "cmd-cc-a",
      idempotencyKey: "cc-a",
      payload: [{ command: "set", path: "/body/paragraph[1]", props: { text: "Variant A" } }]
    });
    await ws.plugin.executeAgentMutation(taskB, {
      commandId: "cmd-cc-b",
      idempotencyKey: "cc-b",
      payload: [{ command: "set", path: "/body/paragraph[1]", props: { text: "Variant B" } }]
    });
    for (const task of [taskA, taskB]) {
      await ws.plugin.flushAgentCandidate(task);
      await ws.plugin.verifyAgentCandidate(task);
      await ws.plugin.finalizeAgentTask(task);
    }
    // Snapshot hashes BEFORE accepting (the winner's candidate row is deleted on accept).
    const hashA = ws.plugin.service.candidates.require(taskA.candidateId).currentHash!;
    const hashB = ws.plugin.service.candidates.require(taskB.candidateId).currentHash!;

    // Fire both accepts concurrently — serialization must let exactly one land.
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

