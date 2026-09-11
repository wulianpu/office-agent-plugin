/**
 * Lease & fencing integration (§38–§40, INV-02/03): single writer per session,
 * monotonic fencing tokens, stale writers rejected even when they return late.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { newCommandId } from "../../src/support/ids.js";
import { openWorkspace, writeDocxFixture } from "../helpers/fixtures.js";
import { createOfficeCliFixture } from "../helpers/officecli-fixture.js";

let ws: Awaited<ReturnType<typeof openWorkspace>>;
let fixture: Awaited<ReturnType<typeof createOfficeCliFixture>>;
let docxPath: string;

beforeAll(async () => {
  fixture = await createOfficeCliFixture();
  ws = await openWorkspace();
  if (fixture.available) {
    docxPath = await fixture.docx(ws.root, "lease.docx");
  } else {
    docxPath = `${ws.root}/lease.docx`;
    await writeDocxFixture(docxPath, ["lease fixture"]);
  }
});

afterAll(async () => {
  await ws?.cleanup().catch(() => undefined);
});

describe("WriterLease & fencing (§38–§40, INV-02/03)", () => {
  it("INV-02: a second writer is rejected while a lease is held", async () => {
    const ref = await ws.plugin.registerArtifact(docxPath);
    const session = await ws.plugin.openSession(ref);
    const { lease } = await ws.plugin.beginEdit(session.sessionId);

    // Agent writer blocked while the human holds the lease.
    await expect(
      ws.plugin.beginAgentTask(session.sessionId, { intent: "x", destructiveAllowed: false })
    ).rejects.toMatchObject({ code: "lease-held" });

    // A second human promotion is blocked as well.
    await expect(ws.plugin.beginEdit(session.sessionId)).rejects.toMatchObject({ code: "lease-held" });
    void lease;
    await ws.plugin.endEdit(session.sessionId);
    await ws.plugin.closeSession(session.sessionId);
  });

  it("INV-03: fencing tokens are monotonic; stale tokens are rejected", async () => {
    const ref = await ws.plugin.registerArtifact(docxPath);
    const session = await ws.plugin.openSession(ref);

    const first = await ws.plugin.beginEdit(session.sessionId);
    const firstToken = first.lease.fencingToken;
    await ws.plugin.endEdit(session.sessionId);

    const second = await ws.plugin.beginEdit(session.sessionId);
    expect(second.lease.fencingToken).toBe(firstToken + 1n);
    await ws.plugin.endEdit(session.sessionId);

    // Simulate a stale writer: a late mutation carrying the first token.
    // Direct lease-manager validation exercises the fence check.
    const task = await ws.plugin.beginAgentTask(session.sessionId, {
      intent: "fence probe",
      destructiveAllowed: false
    });
    const liveLease = ws.plugin.service.leases.activeLease(session.sessionId)!;
    expect(() =>
      ws.plugin.service.leases.validate(
        { ...liveLease, fencingToken: firstToken },
        firstToken
      )
    ).toThrow(/fenced/);
    // The current token validates fine.
    expect(() => ws.plugin.service.leases.validate(liveLease, task.fencingToken)).not.toThrow();
    await ws.plugin.finalizeAgentTask(task);
    await ws.plugin.rejectCandidate(session.sessionId);
    await ws.plugin.closeSession(session.sessionId);
  });

  it("agent mutation with a stale token is rejected end-to-end", async () => {
    const ref = await ws.plugin.registerArtifact(docxPath);
    const session = await ws.plugin.openSession(ref);
    const task = await ws.plugin.beginAgentTask(session.sessionId, {
      intent: "stale end-to-end",
      destructiveAllowed: false
    });
    // Bump the fence: force-release and re-acquire a newer lease.
    ws.plugin.service.leases.releaseForSession(session.sessionId);
    const newLease = ws.plugin.service.leases.acquire(session.sessionId, session.sessionEpoch, {
      sessionId: session.sessionId,
      owner: "agent",
      backend: "officecli",
      baseRevisionId: session.committedRevision.revisionId
    });
    expect(newLease.fencingToken).toBeGreaterThan(task.fencingToken);

    await expect(
      ws.plugin.executeAgentMutation(task, {
        commandId: newCommandId(),
        idempotencyKey: "stale-1",
        payload: [{ command: "set", path: "/body/paragraph[1]", props: { text: "stale" } }]
      })
    ).rejects.toMatchObject({ name: "FencedError" });

    ws.plugin.service.leases.release(newLease.leaseId);
    await ws.plugin.rejectCandidate(session.sessionId);
    await ws.plugin.closeSession(session.sessionId);
  });
});
