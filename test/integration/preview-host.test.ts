/**
 * Localhost host end-to-end (§P7 renderer container): the preview server
 * serves the gallery, slide SVGs, rasterized PNGs and the sidecar-backed xlsx
 * window over HTTP — with no physical path leakage (INV-12).
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { join } from "node:path";
import { PreviewHost } from "../../src/host/preview-server.js";
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

let host: PreviewHost;
let base: string;
let pptxRef: string;
let xlsxRef: string;


afterAll(async () => {
  await host?.stop().catch(() => undefined);
  await ws?.cleanup().catch(() => undefined);
});

describe.skipIf(!engineUp)("localhost preview host", () => {
  beforeAll(async () => {
    fixture = await createOfficeCliFixture();
    ws = await openWorkspace();
    const pptx = await fixture.pptx(ws.root, "host.pptx");
    pptxRef = await ws.plugin.registerArtifact(pptx);
    const xlsx = await fixture.xlsx(ws.root, "host.xlsx");
    xlsxRef = await ws.plugin.registerArtifact(xlsx);
    host = new PreviewHost(ws.plugin);
    const port = await host.start();
    base = `http://127.0.0.1:${port}`;
  });
  it("serves the gallery page", async () => {
    const res = await fetch(base + "/");
    expect(res.status).toBe(200);
    expect(res.headers.get("content-type")).toContain("text/html");
    const page = await res.text();
    expect(page).toContain("Office Plugin");
  });

  it("lists registered artifacts via API (no paths)", async () => {
    const res = await fetch(base + "/api/artifacts");
    const body = (await res.json()) as { artifacts: Array<{ artifactRef: string; format: string }> };
    expect(body.artifacts.length).toBeGreaterThanOrEqual(2);
    expect(JSON.stringify(body)).not.toContain(ws.root);
    expect(JSON.stringify(body)).not.toContain(".pptx\""); // only refs + formats
  });

  it("serves slide SVG with correct content type", async () => {
    const res = await fetch(base + `/svg/${pptxRef}/0`);
    expect(res.status).toBe(200);
    expect(res.headers.get("content-type")).toBe("image/svg+xml");
    const svg = await res.text();
    expect(svg).toMatch(/^<svg xmlns=/);
  });

  it("rasterizes slides to real PNGs on demand", async () => {
    const res = await fetch(base + `/png/${pptxRef}/0`);
    expect(res.status).toBe(200);
    expect(res.headers.get("content-type")).toBe("image/png");
    const png = Buffer.from(await res.arrayBuffer());
    // PNG magic + IHDR.
    expect(png[0]).toBe(0x89);
    expect(png.toString("latin1", 1, 4)).toBe("PNG");
    expect(png.length).toBeGreaterThan(2000);
  });

  it("serves the sidecar-backed xlsx preview window via API", async () => {
    const res = await fetch(base + `/api/preview/${xlsxRef}`);
    const body = (await res.json()) as {
      format: string;
      outline: { kind: "xlsx"; sheets: Array<{ name: string; window: string[][] }> };
    };
    expect(body.format).toBe("xlsx");
    expect(body.outline.kind).toBe("xlsx");
    expect(body.outline.sheets.length).toBeGreaterThan(0);
    expect(JSON.stringify(body)).not.toContain(ws.root);
  });

  it("404s unknown routes and slide indices", async () => {
    expect((await fetch(base + "/nope")).status).toBe(404);
    expect((await fetch(base + `/svg/${pptxRef}/999`)).status).toBe(404);
  });
});

describe.skipIf(!engineUp)("L5 pixel rendering (verification upgrade)", () => {
  it("changed-scope pptx slides rasterize to non-trivial pixels", async () => {
    const session = await ws.plugin.openSession(pptxRef);
    const task = await ws.plugin.beginAgentTask(session.sessionId, {
      intent: "L5 pixel probe",
      destructiveAllowed: false
    });
    await ws.plugin.executeAgentMutation(task, {
      commandId: "cmd-l5",
      idempotencyKey: "l5-1",
      payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Pixel Verified" } }]
    });
    await ws.plugin.flushAgentCandidate(task);
    const report = await ws.plugin.verifyAgentCandidate(task, ["/slide[1]"]);
    expect(report.visual.status).toBe("pass");
    expect(report.confidence).toBe("visual"); // pixel path unlocks the ladder top for engine+diff passes
    await ws.plugin.finalizeAgentTask(task);
    await ws.plugin.rejectCandidate(session.sessionId);
    await ws.plugin.closeSession(session.sessionId);
  });
});
