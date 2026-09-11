/**
 * MCP tool surface integration (§57–§62, INV-12): tools operate through
 * capabilities; results are sanitized — physical paths never leak to agents.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
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
let sessionId: string;


afterAll(async () => {
  await ws?.plugin.closeSession(sessionId).catch(() => undefined);
  await ws?.cleanup().catch(() => undefined);
});

describe.skipIf(!engineUp)("MCP tools (§57–§62, INV-12)", () => {
  beforeAll(async () => {
    fixture = await createOfficeCliFixture();

    ws = await openWorkspace();
    pptxPath = await fixture.pptx(ws.root, "mcp.pptx");
    const ref = await ws.plugin.registerArtifact(pptxPath);
    const session = await ws.plugin.openSession(ref);
    sessionId = session.sessionId;
  });
  it("office.capabilities reports the honest degradation matrix", () => {
    const caps = ws.plugin.mcpTools.capabilities();
    expect(caps.offline).toBe(true);
    const slides = caps.capabilities.find((c) => c.format === "pptx")!;
    const sheets = caps.capabilities.find((c) => c.format === "xlsx")!;
    expect(slides.editor.engine).toBe("genoffice"); // vendored engine (§146)
    expect(sheets.editor.engine).toMatch(/^(genoffice-sheets|basic)$/); // §30 sidecar when built
    expect(slides.agent.status).toBe("available");
  });

  it("office.inspect returns a logical node view with no physical path (INV-12)", async () => {
    const result = await ws.plugin.mcpTools.inspect({ sessionId, path: "/" });
    const json = JSON.stringify(result);
    expect(json).not.toContain(ws.root);
    expect(json).not.toContain(pptxPath);
    expect(result.format).toBe("pptx");
  });

  it("office.query finds shapes by selector", async () => {
    const result = await ws.plugin.mcpTools.query({ sessionId, selector: "shape" });
    expect(result.matches.length).toBeGreaterThan(0);
    expect(JSON.stringify(result)).not.toContain(ws.root);
  });

  it("office.render returns bounded outline sections", async () => {
    const result = await ws.plugin.mcpTools.render({ sessionId, mode: "outline" });
    expect(result.sections.length).toBeGreaterThan(0);
    expect(result.sections[0]?.label).toMatch(/^slide/);
  });

  it("office.edit runs an idempotent scoped mutation through the candidate", async () => {
    const result = await ws.plugin.mcpTools.edit({
      sessionId,
      intent: "mcp: retitle",
      idempotencyKey: "mcp-edit-1",
      scope: { allowedTargets: ["/slide[1]"], destructiveAllowed: false },
      items: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "MCP Title" } }]
    });
    expect(result.receipts).toHaveLength(1);
    expect(result.taskContext.candidateId).toBeTruthy();
    expect(JSON.stringify(result)).not.toContain(ws.root);

    // Cleanup: flush + finalize + reject (review declined).
    const task = ws.plugin.service.agent.taskForCandidate(result.taskContext.candidateId)!;
    await ws.plugin.flushAgentCandidate(task);
    await ws.plugin.finalizeAgentTask(task);
    await ws.plugin.rejectCandidate(sessionId);
  });

  it("office.edit outside the declared scope is policy-denied (§60)", async () => {
    const task = await ws.plugin.beginAgentTask(sessionId, {
      intent: "scope violation probe",
      destructiveAllowed: false,
      allowedTargets: ["/slide[2]"]
    });
    await expect(
      ws.plugin.executeAgentMutation(task, {
        commandId: newCommandId(),
        idempotencyKey: "mcp-scope-1",
        payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "nope" } }]
      })
    ).rejects.toMatchObject({ code: "policy-denied" });
    await ws.plugin.finalizeAgentTask(task);
    await ws.plugin.rejectCandidate(sessionId);
  });
});
