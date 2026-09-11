/**
 * Headless visual rendering + Rust sidecar integration (§P7, §30):
 * - pptx previews carry standalone SVG slides (pptx-render → RenderTree → SVG)
 * - XLSX previews read bounded windows through the compiled Rust sidecar
 * - capability matrix reports the upgraded engines honestly
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { existsSync } from "node:fs";
import { join } from "node:path";
import { openWorkspace } from "../helpers/fixtures.js";
import { createOfficeCliFixture } from "../helpers/officecli-fixture.js";

let ws: Awaited<ReturnType<typeof openWorkspace>>;
let fixture: Awaited<ReturnType<typeof createOfficeCliFixture>>;

/** The large corpus file is the sidecar's meaningful workload; build on demand. */
async function ensureLargeCorpus(): Promise<string> {
  const path = join(process.cwd(), ".corpus", "xlsx-large.xlsx");
  if (existsSync(path)) return path;
  // @ts-expect-error — plain .mjs tool script without declarations
  const { generateCorpus } = await import("../../tools/corpus/generate.mjs");
  await generateCorpus(".corpus");
  return path;
}

beforeAll(async () => {
  fixture = await createOfficeCliFixture();
  ws = await openWorkspace();
});

afterAll(async () => {
  await ws?.cleanup().catch(() => undefined);
});

describe("pptx headless visual rendering (§P7)", () => {
  it("previews embed standalone SVG slides containing laid-out text", async () => {
    const pptx = await fixture.pptx(ws.root, "visual.pptx");
    const ref = await ws.plugin.registerArtifact(pptx);
    const result = await ws.plugin.preview({ artifactRef: ref, priority: "visible" });

    expect(result.model.svgSlides).toBeDefined();
    const svgs = result.model.svgSlides!;
    expect(svgs.length).toBeGreaterThan(0);
    expect(svgs.length).toBeLessThanOrEqual(6); // bounded visual window
    const first = svgs[0]!;
    expect(first).toMatch(/^<svg xmlns=/);
    expect(first).toContain("<text");
    expect(first).toContain("Original"); // laid-out glyph run from the engine
  });

  it("SVG geometry carries resolved pixel boxes, not raw OOXML", async () => {
    const pptx = await fixture.pptx(ws.root, "visual2.pptx");
    const ref = await ws.plugin.registerArtifact(pptx);
    const result = await ws.plugin.preview({ artifactRef: ref, priority: "visible" });
    const first = result.model.svgSlides?.[0]!;
    expect(first).toMatch(/width="\d+"/);
    expect(first).toMatch(/height="\d+"/);
    // Background rect covers the full slide.
    expect(first).toContain("<rect");
  });
});

describe("xlsx Rust sidecar (§30)", () => {
  it("sidecar binary is built and reachable", () => {
    expect(ws.plugin.service.isXlsxSidecarAvailable()).toBe(true);
  });

  it("bounded windows read real values from a 100k-row workbook", async () => {
    const large = await ensureLargeCorpus();
    const ref = await ws.plugin.registerArtifact(large);
    const result = await ws.plugin.preview({ artifactRef: ref, priority: "visible" });
    expect(result.model.outline.kind).toBe("xlsx");
    const sheets = result.model.outline as { kind: "xlsx"; sheets: Array<{ name: string; rowCount: number; window: string[][] }> };
    expect(sheets.sheets[0]?.name).toBe("Big");
    expect(sheets.sheets[0]?.rowCount).toBe(100_000);
    // r*c % 9973 from the corpus generator: row1 = 1,2,3…
    expect(sheets.sheets[0]?.window[0]?.slice(0, 3)).toEqual(["1", "2", "3"]);
    expect(sheets.sheets[0]?.window[1]?.slice(0, 3)).toEqual(["2", "4", "6"]);
  });

  it("capability matrix: xlsx editor available via the Rust sidecar", () => {
    const caps = ws.plugin.mcpTools.capabilities();
    const xlsx = caps.capabilities.find((c) => c.format === "xlsx")!;
    expect(xlsx.editor.status).toBe("available");
    expect(xlsx.editor.engine).toBe("genoffice-sheets");
  });
});
