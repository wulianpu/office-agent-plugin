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

async function waitFor(condition: () => boolean, timeoutMs = 10_000): Promise<boolean> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (condition()) return true;
    await new Promise((resolve) => setTimeout(resolve, 300));
  }
  return condition();
}

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
    const leaseCountBefore = ws.plugin.service.registry.leaseCount();
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
    // P0-2: the editor's ArtifactLease must be released with the editor —
    // registry lease count returns to its pre-edit baseline.
    expect(ws.plugin.service.registry.leaseCount()).toBe(leaseCountBefore);
    await ws.plugin.closeSession(session.sessionId);
  });

  it("external mutation during promotion marks the session conflicting (§13, §73)", async () => {
    const ref = await ws.plugin.registerArtifact(docxPath);
    const session = await ws.plugin.openSession(ref);
    // Mutate the source underneath before promotion.
    const { appendFile } = await import("node:fs/promises");
    const handle = await appendFile(docxPath, "x");
    void handle;
    // The §73 watcher flags the session conflicted (debounced), which gates
    // promotion; the hash-race window itself is covered by the strong-hash
    // comparison when the watcher loses the race.
    const conflicted = await waitFor(
      () => ws.plugin.service.getSession(session.sessionId)?.lifecycle === "conflict"
    );
    expect(conflicted).toBe(true);
    await expect(ws.plugin.beginEdit(session.sessionId)).rejects.toMatchObject({
      code: "recovery-required"
    });
    void sha256File;
    await ws.plugin.closeSession(session.sessionId);
  }, 20_000);

  it("P1-high (#8): XLSX beginEdit no longer fails on the missing full runtime — and leaves NO writer state on any late failure", async () => {
    const { writeXlsxFixture } = await import("../helpers/fixtures.js");
    const xlsxPath = ws.root + "/edit.xlsx";
    await writeXlsxFixture(xlsxPath, [{ name: "Data", rows: [["h1", "h2"]] }]);
    const ref = await ws.plugin.registerArtifact(xlsxPath);
    const session = await ws.plugin.openSession(ref);
    await ws.plugin.service.sessions.ensureStrongIdentity(session.sessionId);

    // Previously deterministic: human lease taken, then "no full FormatRuntime
    // registered for xlsx" — a lease with no editor. Now the basic editor's
    // capability drives the metadata profile and promotion succeeds.
    const { lease, editor } = await ws.plugin.beginEdit(session.sessionId);
    expect(lease.owner).toBe("human");
    expect(editor.state).toBe("clean");
    await ws.plugin.endEdit(session.sessionId);

    // Late-failure compensation: a candidate gate failure AFTER editor
    // preparation must leave zero writer/editor/lease residue.
    const task = await ws.plugin.beginAgentTask(session.sessionId, {
      intent: "block promotion after prepare",
      destructiveAllowed: false
    });
    await expect(ws.plugin.beginEdit(session.sessionId)).rejects.toMatchObject({
      code: "candidate-conflict"
    });
    // The only remaining writer is the AGENT task's own lease — no HUMAN
    // lease was taken, no editor was bound, no dangling ArtifactLease beyond
    // the agent's own context.
    expect(ws.plugin.service.leases.activeLease(session.sessionId)?.owner).toBe("agent");
    expect(ws.plugin.service.getSession(session.sessionId)?.editor).toBeUndefined();
    expect(ws.plugin.service.editorLeaseCount()).toBe(0); // no dangling editor ArtifactLease
    await ws.plugin.finalizeAgentTask(task);
    await ws.plugin.rejectCandidate(session.sessionId);
    await ws.plugin.closeSession(session.sessionId);
  });

  it("P1 (#8): editor.save() goes through the Runtime save gate — external mutation fails closed, never a fake clean", async () => {
    const ref = await ws.plugin.registerArtifact(docxPath);
    const session = await ws.plugin.openSession(ref);
    const { editor } = await ws.plugin.beginEdit(session.sessionId);
    editor.markDirty?.();

    // External save lands while the editor holds unsaved state: the gate
    // (humanSave) detects the source moved off the committed revision.
    const { appendFile } = await import("node:fs/promises");
    await appendFile(docxPath, "x");
    await expect(editor.save()).rejects.toMatchObject({ code: "source-mutated" });
    expect(editor.state).toBe("error"); // NOT clean — the gate failed

    await ws.plugin.closeSession(session.sessionId);
  });

  it("P1-high (#8 reopen): activateEdit failure disposes the mounted editor; binding-persistence failure compensates the human lease", async () => {
    const ref = await ws.plugin.registerArtifact(docxPath);
    const session = await ws.plugin.openSession(ref);
    await ws.plugin.service.sessions.ensureStrongIdentity(session.sessionId);

    // ── Case 1: a plugin whose activateEdit throws ──
    const originalPlugin = ws.plugin.service.editorHost.pluginFor("docx");
    const failing = {
      ...originalPlugin!,
      async create(context: never) {
        const instance = await originalPlugin!.create(context);
        const wrapped = Object.create(instance) as typeof instance & { disposed?: boolean };
        wrapped.activateEdit = async () => {
          throw new Error("engine activation exploded");
        };
        return wrapped;
      }
    };
    ws.plugin.service.editorHost.register(failing as never);
    try {
      await expect(ws.plugin.beginEdit(session.sessionId)).rejects.toThrow("engine activation exploded");
    } finally {
      ws.plugin.service.editorHost.register(originalPlugin!);
    }
    // Zero residue: no human lease, no editor binding, no editor ArtifactLease.
    expect(ws.plugin.service.leases.activeLease(session.sessionId)).toBeUndefined();
    expect(ws.plugin.service.getSession(session.sessionId)?.editor).toBeUndefined();
    expect(ws.plugin.service.editorLeaseCount()).toBe(0);

    // ── Case 2: post-lease session-binding persistence failure ──
    const repos = ws.plugin.service.repos as unknown as {
      upsertSession: (row: unknown) => void;
    };
    const originalUpsert = repos.upsertSession.bind(repos);
    let failNextBinding = false;
    repos.upsertSession = (row: unknown) => {
      const r = row as { lifecycle?: string };
      // Fail only the editor-binding write (lifecycle ready + after lease).
      if (failNextBinding && r.lifecycle === "ready") {
        failNextBinding = false;
        throw new Error("session persistence fault");
      }
      return originalUpsert(row);
    };
    failNextBinding = true;
    try {
      await expect(ws.plugin.beginEdit(session.sessionId)).rejects.toThrow("session persistence fault");
    } finally {
      repos.upsertSession = originalUpsert;
    }
    // Full compensation: human lease released, writerLease/editor cleared,
    // durable lease.released(reason=promotion-failed) recorded.
    expect(ws.plugin.service.leases.activeLease(session.sessionId)).toBeUndefined();
    const live = ws.plugin.service.getSession(session.sessionId);
    expect(live?.writerLease).toBeUndefined();
    expect(live?.editor).toBeUndefined();
    expect(ws.plugin.service.editorLeaseCount()).toBe(0);
    const events = ws.plugin.service.repos.readEventsSince(session.sessionId, 0n, 100);
    const released = [...events].reverse().find((e) => e.type === "lease.released");
    expect((released?.payload as { reason?: string })?.reason).toBe("promotion-failed");

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
