/**
 * WPS host adapter integration (§86–§88): real KWPP COM renders and the
 * disposable certification-copy workflow. The full verification ladder can
 * reach `consumer-certified` when the host certifies the candidate.
 * Top-level probe drives skipIf — COM single-use servers also make the
 * adapter retry concurrent-instantiation rejections internally.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { existsSync } from "node:fs";
import { join } from "node:path";
import { WpsHostAdapter } from "../../src/hosts/wps/wps-adapter.js";
import { openWorkspace } from "../helpers/fixtures.js";
import { createOfficeCliFixture } from "../helpers/officecli-fixture.js";

const adapter = new WpsHostAdapter();
const wpsInstalled = existsSync("C:/Program Files/Kingsoft/WPS Office");
// Stagger the top-level probe so parallel test processes' shared probes
// don't stampede the single-use COM server simultaneously.
await new Promise((r) => setTimeout(r, (process.pid % 5) * 700));
let wpsUp = false;
if (wpsInstalled) {
  // Parallel test processes contend for the single-use COM server; the
  // in-process retry alone is not enough under full-suite fan-out.
  for (let attempt = 0; attempt < 5 && !wpsUp; attempt++) {
    wpsUp = await adapter.probe().catch(() => false);
    if (!wpsUp) await new Promise((r) => setTimeout(r, 3000));
  }
}

let ws: Awaited<ReturnType<typeof openWorkspace>>;
let fixture: Awaited<ReturnType<typeof createOfficeCliFixture>>;
let pptxPath = "";
let probedPlugin: Awaited<ReturnType<typeof import("../../src/plugin/office-plugin.js").OfficePlugin.create>>;

beforeAll(async () => {
  if (!wpsUp) return;
  fixture = await createOfficeCliFixture();
  // Probe-enabled plugin (openWorkspace skips host probes for speed).
  ws = await openWorkspace();
  void ws;
  const { OfficePlugin } = await import("../../src/plugin/office-plugin.js");
  probedPlugin = await OfficePlugin.create({ workspaceRoot: join(ws.root, "probe-rt") });
  await probedPlugin.service.wpsProbePromise;
  pptxPath = await fixture.pptx(ws.root, "wps-host.pptx");
});

afterAll(async () => {
  await probedPlugin?.dispose().catch(() => undefined);
  await ws?.cleanup().catch(() => undefined);
});

const matrixPlugin = () => probedPlugin;

describe.skipIf(!wpsUp)("WPS host adapter (§86–§88)", () => {
  it("probe succeeds", () => {
    expect(wpsUp).toBe(true);
  });

  it("renders real slides through the host (§86 host render)", async () => {
    const result = await adapter.renderSlides(pptxPath, [1], { width: 960, height: 540 });
    expect(result.slideCount).toBeGreaterThanOrEqual(1);
    expect(result.pngs).toHaveLength(1);
    for (const png of result.pngs) {
      expect(png[0]).toBe(0x89);
      expect(png.toString("latin1", 1, 4)).toBe("PNG");
      expect(png.length).toBeGreaterThan(1000);
    }
  }, 240_000);

  it("certifies a disposable copy without repair (§88)", async () => {
    const result = await adapter.certifyCopy(pptxPath);
    expect(result.status).toBe("pass");
    expect(result.repaired).toBe(false);
    expect(result.renderedPng?.length ?? 0).toBeGreaterThan(1000);
  }, 240_000);

  it("full verification reaches consumer-certified with the host", async () => {
    // Runs on the probe-enabled plugin so L7 certification is gated in.
    const plugin = matrixPlugin();
    const ref = await plugin.registerArtifact(pptxPath);
    const session = await plugin.openSession(ref);
    const task = await plugin.beginAgentTask(session.sessionId, {
      intent: "L7 certify probe",
      destructiveAllowed: false
    });
    await plugin.executeAgentMutation(task, {
      commandId: "cmd-l7",
      idempotencyKey: "l7-1",
      payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Host Certified" } }]
    });
    await plugin.flushAgentCandidate(task);
    const report = await plugin.verifyAgentCandidate(task, ["/slide[1]"]);
    expect(report.confidence).toBe("consumer-certified");
    expect(report.visual.issues.some((i) => i.code === "host-certified")).toBe(true);
    await plugin.finalizeAgentTask(task);
    await plugin.rejectCandidate(session.sessionId);
    await plugin.closeSession(session.sessionId);
  }, 300_000);

  it("capability matrix: wps host adapter available", () => {
    const caps = matrixPlugin().mcpTools.capabilities();
    const wps = caps.hostAdapters.find((h) => h.id === "wps")!;
    expect(wps.status.status).toBe("available");
    expect(wps.status.engine).toBe("KWPP-COM");
    const ppt = caps.hostAdapters.find((h) => h.id === "powerpoint")!;
    expect(ppt.status.status).toBe("unavailable");
  });
});
