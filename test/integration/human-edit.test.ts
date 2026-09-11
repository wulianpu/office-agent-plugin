/**
 * Human edit path (§12–§13, §51, §69–§72): read-only open → edit promotion →
 * human save → candidate review lifecycle. Uses hermetic fixtures.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { sha256File } from "../../src/support/fsx.js";
import { openWorkspace, writeDocxFixture } from "../helpers/fixtures.js";

let ws: Awaited<ReturnType<typeof openWorkspace>>;
let docxPath: string;

beforeAll(async () => {
  ws = await openWorkspace();
  docxPath = `${ws.root}/human.docx`;
  await writeDocxFixture(docxPath, ["Human heading", "Human body"]);
});

afterAll(async () => {
  await ws?.cleanup().catch(() => undefined);
});

describe("Human edit path (§12–§13)", () => {
  it("promotes a read-only session to an editor on first mutation", async () => {
    const ref = await ws.plugin.registerArtifact(docxPath);
    const session = await ws.plugin.openSession(ref);
    expect(session.lifecycle).toBe("ready");
    expect(session.writerLease).toBeUndefined();

    // A competing candidate blocks promotion (§13).
    const task = await ws.plugin.beginAgentTask(session.sessionId, {
      intent: "competing candidate",
      destructiveAllowed: false
    });
    await expect(ws.plugin.beginEdit(session.sessionId)).rejects.toMatchObject({
      code: "candidate-conflict"
    });
    await ws.plugin.finalizeAgentTask(task);
    await ws.plugin.rejectCandidate(session.sessionId);

    // With the candidate cleared, promotion succeeds with a human lease.
    const { lease, editor } = await ws.plugin.beginEdit(session.sessionId, {
      location: { block: 1 }
    });
    expect(lease.owner).toBe("human");
    expect(editor.state).toBe("clean");
    expect(ws.plugin.service.leases.activeLease(session.sessionId)?.owner).toBe("human");

    // No-op save on an unchanged source keeps the revision.
    const saved = await ws.plugin.humanSave(session.sessionId);
    expect(saved.unchanged).toBe(true);

    await ws.plugin.endEdit(session.sessionId);
    expect(ws.plugin.service.leases.activeLease(session.sessionId)).toBeUndefined();
    await ws.plugin.closeSession(session.sessionId);
  });

  it("external mutation during promotion marks the session conflicting (§13)", async () => {
    const ref = await ws.plugin.registerArtifact(docxPath);
    const session = await ws.plugin.openSession(ref);
    // Mutate the source underneath before promotion.
    const { appendFile } = await import("node:fs/promises");
    const handle = await appendFile(docxPath, "x");
    void handle;
    await expect(ws.plugin.beginEdit(session.sessionId)).rejects.toMatchObject({
      code: "source-mutated"
    });
    expect(ws.plugin.service.getSession(session.sessionId)?.lifecycle).toBe("conflict");
    void sha256File;
    await ws.plugin.closeSession(session.sessionId);
  });

  it("the editor contract lifecycle holds (mount/save/suspend/resume/dispose)", async () => {
    const ref = await ws.plugin.registerArtifact(docxPath);
    const session = await ws.plugin.openSession(ref);
    const { editor } = await ws.plugin.beginEdit(session.sessionId);
    editor.markDirty?.();
    expect(editor.state).toBe("dirty");
    await editor.save(); // basic editor save path
    expect(editor.state).toBe("clean");
    await editor.suspend();
    await editor.resume();
    await editor.dispose();
    await ws.plugin.endEdit(session.sessionId);
    await ws.plugin.closeSession(session.sessionId);
  });
});
