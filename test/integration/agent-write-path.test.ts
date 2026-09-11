/**
 * Agent write path integration (§158): the full Candidate workflow against a
 * real OfficeCLI engine — task → mutate → flush barrier → verify → handoff →
 * accept with atomic commit. Covers INV-01/04/05/06/07/10/11.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { sha256File } from "../../src/support/fsx.js";
import { newCommandId } from "../../src/support/ids.js";
import { openWorkspace } from "../helpers/fixtures.js";
import { createOfficeCliFixture } from "../helpers/officecli-fixture.js";

let ws: Awaited<ReturnType<typeof openWorkspace>>;
let fixture: Awaited<ReturnType<typeof createOfficeCliFixture>>;
/** CI without the OfficeCLI engine: the whole suite skips cleanly. */
const engineUp = await import("../../src/agent/officecli/officecli-adapter.js")
  .then(async (m) => {
    const probe = new m.OfficeCliAdapter();
    return probe.version_().then(() => true).catch(() => false);
  })
  .catch(() => false);

let pptxPath: string;

beforeAll(async () => {
  fixture = await createOfficeCliFixture();

  ws = await openWorkspace();
  pptxPath = await fixture.pptx(ws.root);
});

afterAll(async () => {
  await ws?.cleanup().catch(() => undefined);
});

describe.skipIf(!engineUp)("Agent write path (§158)", () => {
  it("runs the full candidate workflow: mutate → flush → verify → accept", async () => {
    const ref = await ws.plugin.registerArtifact(pptxPath);
    const session = await ws.plugin.openSession(ref);
    // P0-6: open returns immediately (optimistic identity); writer paths
    // observe the strong revision via ensureStrongIdentity.
    const strong = await ws.plugin.service.sessions.ensureStrongIdentity(session.sessionId);
    const sourceHashBefore = await sha256File(pptxPath);
    expect(strong.contentHash).toBe(sourceHashBefore);

    const task = await ws.plugin.beginAgentTask(session.sessionId, {
      intent: "retitle slide 1",
      destructiveAllowed: false,
      allowedTargets: ["/slide[1]"]
    });
    expect(task.fencingToken).toBeGreaterThan(0n);

    // INV-01: the agent never touches the committed source.
    const candidate = ws.plugin.service.candidates.require(task.candidateId);
    const stagingPath = ws.plugin.service.store.resolvePath(candidate.artifactRef);
    expect(stagingPath).not.toBe(pptxPath);
    // INV-05: candidate bound to base revision + hash.
    expect(candidate.baseRevisionId).toBe(session.committedRevision.revisionId);
    expect(candidate.baseHash).toBe(sourceHashBefore);

    await ws.plugin.executeAgentMutation(task, {
      commandId: newCommandId(),
      idempotencyKey: "agent-full-1",
      payload: [
        { command: "set", path: "/slide[1]/shape[1]", props: { text: "Retitled by Agent" } }
      ]
    });

    // INV-06: resident mutation lives in memory; the candidate file on disk
    // still holds base bytes until the explicit flush barrier.
    const stagingBeforeFlush = await sha256File(stagingPath);
    expect(stagingBeforeFlush).toBe(sourceHashBefore);

    const flushedHash = await ws.plugin.flushAgentCandidate(task);
    expect(flushedHash).not.toBe(sourceHashBefore);
    // INV-01 continues: source untouched.
    expect(await sha256File(pptxPath)).toBe(sourceHashBefore);

    const report = await ws.plugin.verifyAgentCandidate(task, ["/slide[1]"]);
    expect(["structural", "engine", "visual"]).toContain(report.confidence);
    expect(report.contentHash).toBe(flushedHash);

    // INV-07: handoff closes the resident and releases the writer.
    await ws.plugin.finalizeAgentTask(task);
    expect(ws.plugin.service.residentPool.size()).toBe(0);
    expect(ws.plugin.service.leases.activeLease(session.sessionId)).toBeUndefined();

    const accepted = await ws.plugin.acceptCandidate(session.sessionId);
    expect(accepted.contextPromoted).toBe(true);
    const newHash = await sha256File(pptxPath);
    expect(newHash).toBe(flushedHash);

    const updated = ws.plugin.service.getSession(session.sessionId)!;
    expect(updated.committedRevision.sequence).toBe(2);
    expect(updated.committedRevision.origin).toBe("agent");
    await ws.plugin.closeSession(session.sessionId);
  });

  it("P1d: idempotency is candidate-scoped; same key + different payload conflicts (§61)", async () => {
    const ref = await ws.plugin.registerArtifact(pptxPath);
    const session = await ws.plugin.openSession(ref);

    // Task 1 uses the key with payload A.
    const task1 = await ws.plugin.beginAgentTask(session.sessionId, { intent: "t1", destructiveAllowed: false });
    await ws.plugin.executeAgentMutation(task1, {
      commandId: "cmd-a",
      idempotencyKey: "shared-key",
      payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Payload A" } }]
    });

    // Same key + DIFFERENT payload inside the same candidate → conflict.
    await expect(
      ws.plugin.executeAgentMutation(task1, {
        commandId: "cmd-b",
        idempotencyKey: "shared-key",
        payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Payload B" } }]
      })
    ).rejects.toMatchObject({ code: "idempotency-conflict" });

    await ws.plugin.finalizeAgentTask(task1);
    await ws.plugin.rejectCandidate(session.sessionId);

    // Task 2 (new candidate) MAY reuse the key — scope is per candidate.
    const task2 = await ws.plugin.beginAgentTask(session.sessionId, { intent: "t2", destructiveAllowed: false });
    await expect(
      ws.plugin.executeAgentMutation(task2, {
        commandId: "cmd-c",
        idempotencyKey: "shared-key",
        payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Payload C" } }]
      })
    ).resolves.toBeTruthy();
    await ws.plugin.finalizeAgentTask(task2);
    await ws.plugin.rejectCandidate(session.sessionId);
    await ws.plugin.closeSession(session.sessionId);
  });

  it("INV-04: replaying the same idempotencyKey returns the stored receipt without re-executing", async () => {
    const ref = await ws.plugin.registerArtifact(pptxPath);
    const session = await ws.plugin.openSession(ref);
    const task = await ws.plugin.beginAgentTask(session.sessionId, {
      intent: "idempotency check",
      destructiveAllowed: false
    });

    const command = {
      commandId: newCommandId(),
      idempotencyKey: "idem-1",
      payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Once only" } }]
    };
    const first = await ws.plugin.executeAgentMutation(task, command);
    const stagingPath = ws.plugin.service.store.resolvePath(
      ws.plugin.service.candidates.require(task.candidateId).artifactRef
    );
    // Flush so the disk bytes reflect exactly one mutation.
    await ws.plugin.flushAgentCandidate(task);
    const hashAfterFirst = await sha256File(stagingPath);

    const replay = await ws.plugin.executeAgentMutation(task, {
      ...command,
      commandId: newCommandId() // different command id, same idempotency key
    });
    expect(replay.receiptId).toBe(first.receiptId); // same stored receipt
    const hashAfterReplay = await sha256File(stagingPath);
    expect(hashAfterReplay).toBe(hashAfterFirst); // bytes unchanged — no re-execution

    // Cleanup: reject the candidate.
    await ws.plugin.finalizeAgentTask(task);
    await ws.plugin.rejectCandidate(session.sessionId);
    await ws.plugin.closeSession(session.sessionId);
  });
});
