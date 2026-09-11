/**
 * Candidate safety integration (§67–§72, INV-05/08/09/10/11): verification is
 * hash-bound, human amendment voids it, accept checks source hash and always
 * commits through the AtomicFileCommitter.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { writeFile } from "node:fs/promises";
import { sha256File } from "../../src/support/fsx.js";
import { openWorkspace } from "../helpers/fixtures.js";
import { createOfficeCliFixture } from "../helpers/officecli-fixture.js";

let ws: Awaited<ReturnType<typeof openWorkspace>>;
let fixture: Awaited<ReturnType<typeof createOfficeCliFixture>>;

beforeAll(async () => {
  fixture = await createOfficeCliFixture();
  if (!fixture.available) throw new Error("officecli not available");
  ws = await openWorkspace();
});

afterAll(async () => {
  await ws?.cleanup().catch(() => undefined);
});

describe("Candidate safety (§67–§72)", () => {
  it("INV-08: verification is rejected when hashes do not match the candidate", async () => {
    const pptx = await fixture.pptx(ws.root, "verify.pptx");
    const ref = await ws.plugin.registerArtifact(pptx);
    const session = await ws.plugin.openSession(ref);
    const task = await ws.plugin.beginAgentTask(session.sessionId, {
      intent: "verify binding",
      destructiveAllowed: false
    });
    // No flush yet → candidate has no currentHash → verification cannot bind.
    await expect(ws.plugin.verifyAgentCandidate(task)).rejects.toMatchObject({
      code: "candidate-not-ready"
    });

    // Directly attempting to bind a mismatched report is rejected.
    await expect(
      ws.plugin.service.candidates.publishVerification(task.candidateId, session.sessionEpoch, {
        candidateId: task.candidateId,
        contentHash: "deadbeef".repeat(8),
        structural: { layer: "L1-ooxml-structural", status: "pass", issues: [], durationMs: 1 },
        package: { layer: "L2-package-relationships", status: "pass", issues: [], durationMs: 1 },
        semantic: { layer: "L3-officecli-issues", status: "pass", issues: [], durationMs: 1 },
        visual: { layer: "L5-changed-scope-render", status: "skipped", issues: [], durationMs: 1 },
        confidence: "structural",
        verifiedAt: Date.now()
      })
    ).rejects.toMatchObject({ code: "verification-stale" });

    await ws.plugin.finalizeAgentTask(task);
    await ws.plugin.rejectCandidate(session.sessionId);
    await ws.plugin.closeSession(session.sessionId);
  });

  it("INV-09: human amendment invalidates the bound verification", async () => {
    const pptx = await fixture.pptx(ws.root, "amend.pptx");
    const ref = await ws.plugin.registerArtifact(pptx);
    const session = await ws.plugin.openSession(ref);
    const task = await ws.plugin.beginAgentTask(session.sessionId, {
      intent: "amend flow",
      destructiveAllowed: false
    });
    const flushed = await ws.plugin.flushAgentCandidate(task);
    const report = await ws.plugin.verifyAgentCandidate(task);
    expect(report.contentHash).toBe(flushed);
    await ws.plugin.finalizeAgentTask(task);

    const candidate = ws.plugin.service.candidates.require(task.candidateId);
    expect(candidate.verification?.contentHash).toBe(flushed);

    // Human amends the proposal bytes → verification voided.
    const amended = await ws.plugin.service.candidates.markHumanAmended(
      task.candidateId,
      session.sessionEpoch,
      "a".repeat(64)
    );
    expect(amended.verification).toBeUndefined();
    expect(amended.state).toBe("human-amended");
    // Accept is blocked: verification missing for the amended hash.
    await expect(ws.plugin.acceptCandidate(session.sessionId, task.candidateId)).rejects.toMatchObject({
      code: "verification-stale"
    });

    await ws.plugin.rejectCandidate(session.sessionId);
    await ws.plugin.closeSession(session.sessionId);
  });

  it("INV-10: accept fails when the source mutated past the base hash", async () => {
    const pptx = await fixture.pptx(ws.root, "mutated.pptx");
    const ref = await ws.plugin.registerArtifact(pptx);
    const session = await ws.plugin.openSession(ref);
    const task = await ws.plugin.beginAgentTask(session.sessionId, {
      intent: "mutated source",
      destructiveAllowed: false
    });
    await ws.plugin.flushAgentCandidate(task);
    await ws.plugin.verifyAgentCandidate(task);
    await ws.plugin.finalizeAgentTask(task);

    // External mutation of the source while the proposal waited for review.
    // (Append a byte: hash changes, no engine involvement required.)
    const handle = await (await import("node:fs/promises")).open(pptx, "a");
    await handle.write(Buffer.from([0x20]));
    await handle.close();

    await expect(ws.plugin.acceptCandidate(session.sessionId, task.candidateId)).rejects.toMatchObject({
      code: "source-mutated"
    });

    await ws.plugin.rejectCandidate(session.sessionId);
    await ws.plugin.closeSession(session.sessionId);
  });

  it("INV-11: accept lands through the atomic committer (journal finalized)", async () => {
    const pptx = await fixture.pptx(ws.root, "journal.pptx");
    const ref = await ws.plugin.registerArtifact(pptx);
    const session = await ws.plugin.openSession(ref);
    const task = await ws.plugin.beginAgentTask(session.sessionId, {
      intent: "journal check",
      destructiveAllowed: false
    });
    await ws.plugin.executeAgentMutation(task, {
      commandId: "cmd-journal",
      idempotencyKey: "journal-1",
      payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Journaled" } }]
    });
    await ws.plugin.flushAgentCandidate(task);
    await ws.plugin.verifyAgentCandidate(task);
    await ws.plugin.finalizeAgentTask(task);

    await ws.plugin.acceptCandidate(session.sessionId, task.candidateId);
    const journal = ws.plugin.service.repos.listJournal();
    const last = journal[journal.length - 1]!;
    expect(last.phase).toBe("finalized");
    expect(await sha256File(pptx)).toBe(last.candidateHash);
    await ws.plugin.closeSession(session.sessionId);
  });
});
