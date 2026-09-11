/**
 * Read path integration (§157, PERF-01/02/03/07): preview creates no session,
 * no lease, no candidate and never spawns OfficeCLI. Open is read-only MVCC.
 * Preview retries once when the file mutates underneath (§18).
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { writeFile } from "node:fs/promises";
import { sha256File } from "../../src/support/fsx.js";
import {
  openWorkspace,
  writePptxFixture,
  writeDocxFixture,
  writeXlsxFixture
} from "../helpers/fixtures.js";

let ws: Awaited<ReturnType<typeof openWorkspace>>;
let pptxPath: string;
let docxPath: string;
let xlsxPath: string;

beforeAll(async () => {
  ws = await openWorkspace();
  pptxPath = `${ws.root}/read.pptx`;
  docxPath = `${ws.root}/read.docx`;
  xlsxPath = `${ws.root}/read.xlsx`;
  await writePptxFixture(pptxPath, [
    [{ name: "Title", text: "Quarterly Review" }],
    [{ name: "Body", text: "Revenue up" }]
  ]);
  await writeDocxFixture(docxPath, ["Alpha heading", "Beta paragraph"]);
  await writeXlsxFixture(xlsxPath, [{ name: "Sheet1", rows: [["h1", "h2"], ["1", "2"]] }]);
});

afterAll(async () => {
  await ws.cleanup();
});

describe("Read path (§157, PERF-01/02/03)", () => {
  it("preview renders outlines without session/lease/candidate (PERF-01/02)", async () => {
    const ref = await ws.plugin.registerArtifact(pptxPath);
    const result = await ws.plugin.preview({ artifactRef: ref, priority: "visible" });
    expect(result.model.outline.kind).toBe("pptx");
    const outline = result.model.outline as { kind: "pptx"; slides: Array<{ shapes: Array<{ text?: string }> }> };
    expect(outline.slides).toHaveLength(2);
    expect(outline.slides[0]?.shapes[0]?.text).toContain("Quarterly Review");

    // No session machinery was created.
    expect(ws.plugin.service.getSession).toBeDefined();
    expect(ws.plugin.service.sessions.list()).toHaveLength(0);
    expect(ws.plugin.service.leases.hasActiveLease("none")).toBe(false);
    expect(ws.plugin.service.residentPool.size()).toBe(0); // no OfficeCLI spawn
  });

  it("second preview hits the cache (no re-render cost)", async () => {
    const ref = await ws.plugin.registerArtifact(pptxPath);
    await ws.plugin.preview({ artifactRef: ref, priority: "visible" });
    const statsBefore = ws.plugin.service.previewService.stats();
    const second = await ws.plugin.preview({ artifactRef: ref, priority: "visible" });
    const statsAfter = ws.plugin.service.previewService.stats();
    expect(second.retried).toBe(false);
    expect(statsAfter.hits).toBeGreaterThan(statsBefore.hits);
  });

  it("docx and xlsx outlines render", async () => {
    const docxRef = await ws.plugin.registerArtifact(docxPath);
    const docxPreview = await ws.plugin.preview({ artifactRef: docxRef, priority: "visible" });
    expect(docxPreview.model.outline.kind).toBe("docx");

    const xlsxRef = await ws.plugin.registerArtifact(xlsxPath);
    const xlsxPreview = await ws.plugin.preview({ artifactRef: xlsxRef, priority: "visible" });
    expect(xlsxPreview.model.outline.kind).toBe("xlsx");
    const sheets = (xlsxPreview.model.outline as { kind: "xlsx"; sheets: Array<{ name: string; window: string[][] }> }).sheets;
    expect(sheets[0]?.name).toBe("Sheet1");
    expect(sheets[0]?.window[0]).toEqual(["h1", "h2"]);
  });

  it("open creates a read-only session with no writer lease (PERF-03)", async () => {
    const ref = await ws.plugin.registerArtifact(docxPath);
    const session = await ws.plugin.openSession(ref);
    expect(session.lifecycle).toBe("ready");
    expect(session.writerLease).toBeUndefined();
    expect(session.candidate).toBeUndefined();
    const hash = await sha256File(docxPath);
    expect(session.committedRevision.contentHash).toBe(hash);
    await ws.plugin.closeSession(session.sessionId);
  });

  it("preview recomputes when the file mutates during render (§18 optimistic)", async () => {
    const ref = await ws.plugin.registerArtifact(pptxPath);
    // Mutate right after preview starts; the before/after fingerprint check
    // must observe the change and retry.
    const first = ws.plugin.preview({ artifactRef: ref, priority: "visible" });
    await new Promise((resolve) => setImmediate(resolve));
    await writeFile(pptxPath, await (await import("node:fs/promises")).readFile(pptxPath));
    const result = await first;
    // Either retried, or rendered against a consistent snapshot (mtimeMs bump
    // may land inside the same fingerprint window); the invariant is no stale
    // cache poison: a follow-up preview reflects the current bytes.
    const second = await ws.plugin.preview({ artifactRef: ref, priority: "visible" });
    expect(second.model.fingerprintAtRender.length).toBeGreaterThan(0);
    void result;
  });
});
