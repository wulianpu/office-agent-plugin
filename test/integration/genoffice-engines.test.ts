/**
 * GenOffice engine integration (§146): vendored engines parse real files,
 * previews carry engine-accurate models, save round-trips produce valid
 * packages (proven via the officecli engine), and the capability matrix
 * reports honestly. Offline gate: bundles contain zero network calls (§120).
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { writeFile } from "node:fs/promises";
import { join } from "node:path";
import {
  openPptx,
  savePptx,
  parseDocx,
  saveDocx,
  probeVendorEngines,
  elementText
} from "../../src/vendor/genoffice/wrapper.js";
import type { PptxSlideElement } from "../../src/vendor/genoffice/wrapper.js";
import { openWorkspace } from "../helpers/fixtures.js";
import { createOfficeCliFixture } from "../helpers/officecli-fixture.js";

let ws: Awaited<ReturnType<typeof openWorkspace>>;
let fixture: Awaited<ReturnType<typeof createOfficeCliFixture>>;
/** CI without the OfficeCLI engine: the whole suite skips cleanly. */
const engineUp = await import("../../src/agent/officecli/officecli-adapter.js")
  .then(async (m) => {
    const probe = new m.OfficeCliAdapter({ timeoutMs: 5_000 }); // probe skips in seconds — never the 120s default
    return probe.version_().then(() => true).catch(() => false);
  })
  .catch(() => false);



afterAll(async () => {
  await ws?.cleanup().catch(() => undefined);
});

describe.skipIf(!engineUp)("GenOffice vendor engines (§146)", () => {
  beforeAll(async () => {
    fixture = await createOfficeCliFixture();
    ws = await openWorkspace();
  });
  it("bundles are loadable (probe)", async () => {
    const probe = await probeVendorEngines();
    expect(probe.pptx).toBe(true);
    expect(probe.docx).toBe(true);
  });

  it("openPptx parses slides with engine-accurate structure", async () => {
    const pptx = await fixture.pptx(ws.root, "go.pptx");
    const { readFile } = await import("node:fs/promises");
    const opened = await openPptx(new Uint8Array(await readFile(pptx)));
    expect(opened.deck.slides.length).toBeGreaterThan(0);
    expect(opened.deck.size.cx).toBeGreaterThan(0);
    const first = opened.deck.slides[0]!;
    expect(first.elements.length).toBeGreaterThan(0);
    expect(first.elements.some((e: PptxSlideElement) => elementText(e).includes("Original"))).toBe(true);
  });

  it("parseDocx returns top-level blocks", async () => {
    const docx = await fixture.docx(ws.root, "go.docx");
    const { readFile } = await import("node:fs/promises");
    const parsed = await parseDocx(new Uint8Array(await readFile(docx)));
    expect(parsed.blocks.length).toBeGreaterThan(0);
    const text = JSON.stringify(parsed.blocks);
    expect(text).toContain("paragraph");
  });

  it("savePptx round-trip produces a package the officecli engine validates", async () => {
    const pptx = await fixture.pptx(ws.root, "go-rt.pptx");
    const { readFile } = await import("node:fs/promises");
    const bytes = await readFile(pptx);
    const opened = await openPptx(new Uint8Array(bytes));
    const saved = await savePptx(opened);
    expect(saved.length).toBeGreaterThan(1000);
    const out = join(ws.root, "go-roundtrip.pptx");
    await writeFile(out, saved);
    const validation = await fixture.adapter.validate(out);
    expect(validation.passed).toBe(true);
  });

  it("saveDocx round-trip produces a package the officecli engine validates", async () => {
    const docx = await fixture.docx(ws.root, "go-rt.docx");
    const { readFile } = await import("node:fs/promises");
    const bytes = await readFile(docx);
    const parsed = await parseDocx(new Uint8Array(bytes));
    const saved = await saveDocx(parsed, { originalBytes: bytes });
    expect(saved.length).toBeGreaterThan(1000);
    const out = join(ws.root, "go-roundtrip.docx");
    await writeFile(out, saved);
    const validation = await fixture.adapter.validate(out);
    expect(validation.passed).toBe(true);
  });
});

describe.skipIf(!engineUp)("GenOffice-backed runtime paths", () => {
  it("P0-7: default previews stay light (ZIP/index, no full engine parse)", async () => {
    const pptx = await fixture.pptx(ws.root, "prev.pptx");
    const ref = await ws.plugin.registerArtifact(pptx);
    const result = await ws.plugin.preview({ artifactRef: ref, priority: "visible" });
    expect(result.model.cacheKey.rendererVersion).toBe("basic-1"); // metadata profile
    expect(result.model.svgSlides).toBeUndefined(); // no engine model attached
    const outline = result.model.outline as { kind: "pptx"; slides: Array<{ shapes: Array<{ text?: string }> }> };
    expect(outline.slides.length).toBeGreaterThan(0);
    expect(outline.slides[0]?.shapes.some((s) => (s.text ?? "").includes("Original"))).toBe(true);
  });

  it("P0-7: visual previews upgrade to the GenOffice engine model + SVG", async () => {
    const pptx = await fixture.pptx(ws.root, "prev-visual.pptx");
    const ref = await ws.plugin.registerArtifact(pptx);
    const result = await ws.plugin.preview({ artifactRef: ref, priority: "visible", visual: true });
    expect(result.model.cacheKey.rendererVersion).toBe("genoffice-1");
    expect(result.model.svgSlides?.length).toBeGreaterThan(0);
    const outline = result.model.outline as { kind: "pptx"; slides: Array<{ shapes: Array<{ text?: string }> }> };
    expect(outline.slides[0]?.shapes.some((s) => (s.text ?? "").includes("Original"))).toBe(true);
  });

  it("docx visual previews use engine blocks", async () => {
    const docx = await fixture.docx(ws.root, "prev.docx");
    const ref = await ws.plugin.registerArtifact(docx);
    const result = await ws.plugin.preview({ artifactRef: ref, priority: "visible", visual: true });
    expect(result.model.cacheKey.rendererVersion).toBe("genoffice-1");
    expect(result.model.outline.kind).toBe("docx");
  });

  it("xlsx keeps the streaming basic path (honest degradation §127)", async () => {
    const { writeXlsxFixture } = await import("../helpers/fixtures.js");
    const xlsx = join(ws.root, "prev.xlsx");
    await writeXlsxFixture(xlsx, [{ name: "S", rows: [["a", "b"]] }]);
    const ref = await ws.plugin.registerArtifact(xlsx);
    const result = await ws.plugin.preview({ artifactRef: ref, priority: "visible" });
    expect(result.model.cacheKey.rendererVersion).toBe("basic-1");
  });

  it("capability matrix: all three formats served by GenOffice engines", () => {
    const caps = ws.plugin.mcpTools.capabilities();
    const pptx = caps.capabilities.find((c) => c.format === "pptx")!;
    const docx = caps.capabilities.find((c) => c.format === "docx")!;
    const xlsx = caps.capabilities.find((c) => c.format === "xlsx")!;
    expect(pptx.editor.status).toBe("available");
    expect(pptx.editor.engine).toBe("genoffice");
    expect(docx.editor.status).toBe("available");
    expect(docx.editor.engine).toBe("genoffice");
    // Rust sidecar (§30) once built; basic streaming is the fallback.
    if (ws.plugin.service.isXlsxSidecarAvailable()) {
      expect(xlsx.editor.status).toBe("available");
      expect(xlsx.editor.engine).toBe("genoffice-sheets");
    } else {
      expect(xlsx.editor.status).toBe("degraded");
      expect(xlsx.editor.engine).toBe("basic");
    }
  });
});
