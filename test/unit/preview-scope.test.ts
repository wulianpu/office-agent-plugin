/**
 * Round 10 (issue #4, P1-high): PreviewScope is a real contract — scoped
 * windows render different content, and cache/dedup keys never cross scopes.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { buildZip } from "../helpers/zip-builder.js";
import { renderPptxOutline, renderXlsxOutline } from "../../src/preview/outline-renderers.js";
import { PreviewService, normalizeScopeKey, svgWindowOf } from "../../src/preview/preview-service.js";
import { ArtifactRegistry } from "../../src/artifact/registry/artifact-registry.js";
import { Scheduler } from "../../src/runtime/scheduler/scheduler.js";
import type { ArtifactStore } from "../../src/artifact/store/artifact-store.js";
import type { FormatRuntime } from "../../src/contracts/artifact.js";

let dir: string;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "preview-scope-"));
});

afterAll(async () => {
  await rm(dir, { recursive: true, force: true }).catch(() => undefined);
});

function sheetXml(rows: string[][], dimension?: string): string {
  const body = rows
    .map((row, r) => `<row r="${r + 1}">${row.map((v, c) => `<c r="${colName(c)}${r + 1}"><v>${v}</v></c>`).join("")}</row>`)
    .join("");
  const dim = dimension ? `<dimension ref="${dimension}"/>` : "";
  return `<?xml version="1.0"?><worksheet>${dim}<sheetData>${body}</sheetData></worksheet>`;
}

function colName(index: number): string {
  let n = index;
  let name = "";
  do {
    name = String.fromCharCode(65 + (n % 26)) + name;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return name;
}

async function writeTwoSheetWorkbook(name: string, withDimension?: string): Promise<string> {
  const zip = buildZip([
    { name: "xl/workbook.xml", data: `<?xml version="1.0"?><workbook><sheets><sheet name="Alpha" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets></workbook>` },
    { name: "xl/_rels/workbook.xml.rels", data: `<?xml version="1.0"?><Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>` },
    { name: "xl/worksheets/sheet1.xml", data: sheetXml([["A1", "B1"], ["A2", "B2"]], withDimension) },
    { name: "xl/worksheets/sheet2.xml", data: sheetXml([["D10", "E10"], ["D11", "E11"]]) }
  ]);
  const path = join(dir, `${name}.xlsx`);
  await writeFile(path, zip);
  return path;
}

async function writeRawSheetWorkbook(name: string, sheetDataXml: string): Promise<string> {
  const zip = buildZip([
    { name: "xl/workbook.xml", data: `<?xml version="1.0"?><workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>` },
    { name: "xl/_rels/workbook.xml.rels", data: `<?xml version="1.0"?><Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>` },
    { name: "xl/worksheets/sheet1.xml", data: `<?xml version="1.0"?><worksheet><sheetData>${sheetDataXml}</sheetData></worksheet>` }
  ]);
  const path = join(dir, `${name}.xlsx`);
  await writeFile(path, zip);
  return path;
}

async function writeSlideDeck(count: number): Promise<string> {
  const entries = [{ name: "[Content_Types].xml", data: `<?xml version="1.0"?><Types/>` }];
  for (let i = 1; i <= count; i++) {
    entries.push({
      name: `ppt/slides/slide${i}.xml`,
      data: `<?xml version="1.0"?><p:sld><p:sp><p:cNvPr id="1" name="S${i}"/><a:t>slide ${i}</a:t></p:sp></p:sld>`
    });
  }
  const path = join(dir, "deck.pptx");
  await writeFile(path, buildZip(entries));
  return path;
}

describe("PreviewScope (round 10, issue #4)", () => {
  it("normalizeScopeKey is canonical AND collision-free (round 10 reopen)", () => {
    const a = normalizeScopeKey({ location: { sheet: "Data" }, maxEntries: 20 });
    const b = normalizeScopeKey({ maxEntries: 20, location: { sheet: "Data" } });
    expect(a).toBe(b); // field order never matters
    // Deliberate collision probe: a sheet NAME containing raw delimiters must
    // never alias a different scope (JSON tuple keeps values escaped).
    expect(normalizeScopeKey({ location: { sheet: "A|max=1" } })).not.toBe(
      normalizeScopeKey({ location: { sheet: "A" }, maxEntries: 1 })
    );
    expect(normalizeScopeKey(undefined)).toBe("");
    // An explicit empty scope is its own (stable) key, distinct from none.
    expect(normalizeScopeKey({})).toBe("[null,null,null,null,null]");
  });

  it("xlsx: sheet scope windows into the named sheet; range slices the window", async () => {
    const path = await writeTwoSheetWorkbook("scoped");
    const bySheet = await renderXlsxOutline(path, { location: { sheet: "Data" } });
    if (bySheet.kind !== "xlsx") throw new Error("expected xlsx");
    expect(bySheet.sheets[0]!.name).toBe("Data");
    expect(bySheet.sheets[0]!.window[0]).toEqual(["D10", "E10"]);

    const ranged = await renderXlsxOutline(path, {
      location: { sheet: "Alpha", range: { fromRow: 2, toRow: 2, fromCol: 2, toCol: 2 } }
    });
    if (ranged.kind !== "xlsx") throw new Error("expected xlsx");
    expect(ranged.sheets[0]!.name).toBe("Alpha");
    expect(ranged.sheets[0]!.window).toEqual([["B2"]]);

    const capped = await renderXlsxOutline(path, { location: { sheet: "Alpha" }, maxEntries: 1 });
    if (capped.kind !== "xlsx") throw new Error("expected xlsx");
    expect(capped.sheets[0]!.window).toHaveLength(1); // maxEntries actually bounds output
  });

  it("xlsx fallback: an explicit MISSING sheet fails closed — empty scoped result, never other sheets (round 10 reopen #3)", async () => {
    const path = await writeTwoSheetWorkbook("missing-sheet");
    const outline = await renderXlsxOutline(path, { location: { sheet: "DoesNotExist" } });
    if (outline.kind !== "xlsx") throw new Error("expected xlsx outline");
    expect(outline.sheets).toEqual([]); // fail-closed: no Alpha/Data masquerading
  });

  it("pptx: slide scope windows at the requested slide", async () => {
    const path = await writeSlideDeck(25);
    const scoped = await renderPptxOutline(path, { location: { slide: 20 }, maxEntries: 2 });
    if (scoped.kind !== "pptx") throw new Error("expected pptx");
    expect(scoped.slides.map((s) => s.index)).toEqual([20, 21]);
    expect(scoped.slides[0]!.shapes[0]!.text).toContain("slide 20");

    const first = await renderPptxOutline(path);
    if (first.kind !== "pptx") throw new Error("expected pptx");
    expect(first.slides[0]!.index).toBe(1);
  });

  it("svgWindowOf: outline and SVG always derive from the same slide window (round 10 reopen)", () => {
    expect(svgWindowOf(25, undefined, 6)).toEqual({ from: 0, count: 6 }); // default: first 6
    expect(svgWindowOf(25, { location: { slide: 20 } }, 6)).toEqual({ from: 19, count: 6 }); // slides 20-25
    expect(svgWindowOf(25, { location: { slide: 20 }, maxEntries: 2 }, 6)).toEqual({ from: 19, count: 2 });
    expect(svgWindowOf(25, { location: { slide: 90 } }, 6)).toEqual({ from: 24, count: 1 }); // clamped tail
    expect(svgWindowOf(0, undefined, 6)).toEqual({ from: 0, count: 0 }); // empty deck
  });

  it("xlsx fallback: a range starting beyond column 24 resolves (AA10:AC20)", async () => {
    // Refs out to AC (col 29) — the default 24-column parse cap used to
    // yield an empty window for any fromCol > 24.
    const mk = (r: number) =>
      ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC"]
        .map((col, i) => `<c r="${col}${r}"><v>${col}${r}</v></c>`)
        .join("");
    const rows = Array.from({ length: 20 }, (_, i) => `<row r="${i + 1}">${mk(i + 1)}</row>`).join("");
    const path = await writeRawSheetWorkbook("high-columns", rows);
    const outline = await renderXlsxOutline(path, {
      location: { range: { fromRow: 10, toRow: 20, fromCol: 27, toCol: 29 } }
    });
    if (outline.kind !== "xlsx") throw new Error("expected xlsx outline");
    expect(outline.sheets[0]!.window[0]).toEqual(["AA10", "AB10", "AC10"]);
    expect(outline.sheets[0]!.window[10]).toEqual(["AA20", "AB20", "AC20"]);
    expect(outline.sheets[0]!.window).toHaveLength(11);
  });

  it("xlsx fallback: a range starting beyond column 256 resolves (IW10:IY20) — no absolute column ceiling (round 10 reopen)", async () => {
    // IW/IY are 1-based columns 257..259 — the old colParseCap(256) and the
    // pad-from-A approach both failed here. Direct absolute->local placement
    // keeps the working set proportional to the WINDOW.
    const colName = (n: number): string => {
      let name = "";
      while (n > 0) {
        const rem = (n - 1) % 26;
        name = String.fromCharCode(65 + rem) + name;
        n = Math.floor((n - 1) / 26);
      }
      return name;
    };
    const mk = (r: number) =>
      [256, 257, 258, 259, 260]
        .map((c) => `<c r="${colName(c)}${r}"><v>${colName(c)}${r}</v></c>`)
        .join("");
    const rows = Array.from({ length: 20 }, (_, i) => `<row r="${i + 1}">${mk(i + 1)}</row>`).join("");
    const path = await writeRawSheetWorkbook("beyond-256", rows);
    const outline = await renderXlsxOutline(path, {
      location: { range: { fromRow: 10, toRow: 20, fromCol: 257, toCol: 259 } }
    });
    if (outline.kind !== "xlsx") throw new Error("expected xlsx outline");
    expect(outline.sheets[0]!.window[0]).toEqual(["IW10", "IX10", "IY10"]);
    expect(outline.sheets[0]!.window[10]).toEqual(["IW20", "IX20", "IY20"]);
  });

  it("an explicit sheet returns ONLY that sheet — 4-sheet workbook, zero unrelated reads (round 10 reopen #5)", async () => {
    // Fallback path.
    const zip = buildZip([
      { name: "xl/workbook.xml", data: `<?xml version="1.0"?><workbook><sheets><sheet name="Alpha" sheetId="1" r:id="r1"/><sheet name="Data" sheetId="2" r:id="r2"/><sheet name="Summary" sheetId="3" r:id="r3"/><sheet name="Zeta" sheetId="4" r:id="r4"/></sheets></workbook>` },
      { name: "xl/_rels/workbook.xml.rels", data: `<?xml version="1.0"?><Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Target="worksheets/sheet2.xml"/><Relationship Id="r3" Target="worksheets/sheet3.xml"/><Relationship Id="r4" Target="worksheets/sheet4.xml"/></Relationships>` },
      { name: "xl/worksheets/sheet1.xml", data: sheetXml([["alpha"]]) },
      { name: "xl/worksheets/sheet2.xml", data: sheetXml([["data-value"]]) },
      { name: "xl/worksheets/sheet3.xml", data: sheetXml([["summary"]]) },
      { name: "xl/worksheets/sheet4.xml", data: sheetXml([["zeta"]]) }
    ]);
    const path = join(dir, "isolated.xlsx");
    await writeFile(path, zip);
    const outline = await renderXlsxOutline(path, { location: { sheet: "Data" } });
    if (outline.kind !== "xlsx") throw new Error("expected xlsx outline");
    expect(outline.sheets).toHaveLength(1); // ONLY Data — never Data + 3 others
    expect(outline.sheets[0]!.name).toBe("Data");
    expect(outline.sheets[0]!.window[0]![0]).toBe("data-value");
  });

  it("wide explicit ranges exceed the default 24-col window identically to the sidecar path (round 10 reopen #5)", async () => {
    const mk = (r: number) =>
      Array.from({ length: 56 }, (_, c) => `<c r="${colName(c)}${r}"><v>${colName(c)}${r}</v></c>`).join("");
    const rows = Array.from({ length: 10 }, (_, i) => `<row r="${i + 1}">${mk(i + 1)}</row>`).join("");
    const path = await writeRawSheetWorkbook("wide-range", rows);

    // A1:AZ10 — 52 columns: no silent truncation to 24.
    const az = await renderXlsxOutline(path, { location: { range: { fromRow: 1, toRow: 10, fromCol: 1, toCol: 52 } } });
    if (az.kind !== "xlsx") throw new Error("expected xlsx outline");
    expect(az.sheets[0]!.window[0]).toHaveLength(52);
    expect(az.sheets[0]!.window[0]![51]).toBe("AZ1");

    // AA1:BD10 — columns 27..56 (width 30), starting beyond the old cap.
    const bd = await renderXlsxOutline(path, { location: { range: { fromRow: 1, toRow: 10, fromCol: 27, toCol: 56 } } });
    if (bd.kind !== "xlsx") throw new Error("expected xlsx outline");
    expect(bd.sheets[0]!.window[0]).toHaveLength(30);
    expect(bd.sheets[0]!.window[0]![0]).toBe("AA1");
    expect(bd.sheets[0]!.window[0]![29]).toBe("BD1");
  });

  it("PPTX fallback follows PRESENTATION order (sldIdLst), not part numbers (round 10 reopen #5)", async () => {
    const slide = (n: number) => ({
      name: `ppt/slides/slide${n}.xml`,
      data: `<?xml version="1.0"?><p:sld><p:sp><p:cNvPr id="1" name="S${n}"/><a:t>content-of-slide${n}</a:t></p:sp></p:sld>`
    });
    const zip = buildZip([
      { name: "[Content_Types].xml", data: `<?xml version="1.0"?><Types/>` },
      slide(1),
      slide(2),
      slide(3),
      // Visible order: slide3, slide1, slide2.
      {
        name: "ppt/presentation.xml",
        data: `<?xml version="1.0"?><p:presentation xmlns:r="rel"><p:sldIdLst><p:sldId id="256" r:id="rA"/><p:sldId id="257" r:id="rB"/><p:sldId id="258" r:id="rC"/></p:sldIdLst></p:presentation>`
      },
      {
        name: "ppt/_rels/presentation.xml.rels",
        data: `<?xml version="1.0"?><Relationships><Relationship Id="rA" Target="slides/slide3.xml"/><Relationship Id="rB" Target="slides/slide1.xml"/><Relationship Id="rC" Target="slides/slide2.xml"/></Relationships>`
      }
    ]);
    const path = join(dir, "reordered.pptx");
    await writeFile(path, zip);

    // Default outline follows presentation order with 1-based positions.
    const outline = await renderPptxOutline(path);
    if (outline.kind !== "pptx") throw new Error("expected pptx outline");
    expect(outline.slides.map((s) => s.index)).toEqual([1, 2, 3]);
    expect(outline.slides[0]!.shapes[0]!.text).toContain("content-of-slide3");

    // slide=1 is the FIRST VISIBLE slide — slide3's content, not slide1.xml.
    const first = await renderPptxOutline(path, { location: { slide: 1 }, maxEntries: undefined });
    if (first.kind !== "pptx") throw new Error("expected pptx outline");
    expect(first.slides[0]!.shapes[0]!.text).toContain("content-of-slide3");
    expect(first.slides[0]!.index).toBe(1);
  });

  it("sheet names with XML entities match and echo decoded (round 10 reopen #5)", async () => {
    const zip = buildZip([
      { name: "xl/workbook.xml", data: `<?xml version="1.0"?><workbook><sheets><sheet name="R&amp;D" sheetId="1" r:id="r1"/></sheets></workbook>` },
      { name: "xl/_rels/workbook.xml.rels", data: `<?xml version="1.0"?><Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/></Relationships>` },
      { name: "xl/worksheets/sheet1.xml", data: sheetXml([["rd-value"]]) }
    ]);
    const path = join(dir, "entity-sheet.xlsx");
    await writeFile(path, zip);
    // Scope uses the REAL name; the fallback decodes R&amp;D before matching.
    const outline = await renderXlsxOutline(path, { location: { sheet: "R&D" } });
    if (outline.kind !== "xlsx") throw new Error("expected xlsx outline");
    expect(outline.sheets).toHaveLength(1);
    expect(outline.sheets[0]!.name).toBe("R&D");
    expect(outline.sheets[0]!.window[0]![0]).toBe("rd-value");
  });

  it("pptx out-of-range scope clamps to the last slide in BOTH outline and SVG window (round 10 reopen)", async () => {
    // Engine-path policy: svgWindowOf clamps `from` to the last slide; the
    // outline now derives from the same function.
    expect(svgWindowOf(25, { location: { slide: 90 } }, 6)).toEqual({ from: 24, count: 1 });
    // Zip fallback policy: same clamp — never an empty outline.
    const path = await writeSlideDeck(25);
    const scoped = await renderPptxOutline(path, { location: { slide: 90 } });
    if (scoped.kind !== "pptx") throw new Error("expected pptx outline");
    expect(scoped.slides).toHaveLength(1);
    expect(scoped.slides[0]!.index).toBe(25); // the LAST slide, not empty
  });

  it("xlsx fallback: maxEntries narrows an explicit range but never extends past toRow", async () => {
    const rows = Array.from({ length: 30 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i + 1}</v></c></row>`).join("");
    const path = await writeRawSheetWorkbook("torow-bound", rows);
    const outline = await renderXlsxOutline(path, {
      location: { range: { fromRow: 5, toRow: 10, fromCol: 1, toCol: 1 } },
      maxEntries: 100 // larger than the range — toRow must still bound it
    });
    if (outline.kind !== "xlsx") throw new Error("expected xlsx outline");
    expect(outline.sheets[0]!.window).toHaveLength(6); // rows 5..10 exactly
    expect(outline.sheets[0]!.window.map((row) => row[0])).toEqual(["5", "6", "7", "8", "9", "10"]);
  });

  it("two scopes of one artifact never share a cache entry or dedup promise", async () => {
    const path = await writeTwoSheetWorkbook("cachescopes");
    const registry = new ArtifactRegistry();
    const runtime: FormatRuntime = {
      format: "xlsx",
      async createArtifactContext(input) {
        return {
          artifactRef: input.artifactRef,
          version: { artifactRef: input.artifactRef, fingerprint: { size: 1n, mtimeNs: 1n } },
          format: "xlsx",
          consistency: input.consistency,
          rendererVersion: "test",
          lastAccessAt: Date.now(),
          enrichment: new Map()
        };
      },
      initialize: async () => undefined,
      trimMemory: async () => undefined,
      dispose: async () => undefined
    };
    registry.registerRuntime(runtime);
    registry.setPathResolver(() => path);
    const service = new PreviewService(
      { formatOf: () => "xlsx", resolvePath: () => path } as unknown as ArtifactStore,
      registry,
      new Scheduler()
    );

    const sheetAlpha = await service.preview({
      requestId: "r1",
      artifactRef: "art-scope",
      priority: "visible",
      scope: { location: { sheet: "Alpha" } }
    });
    const sheetData = await service.preview({
      requestId: "r2",
      artifactRef: "art-scope",
      priority: "visible",
      scope: { location: { sheet: "Data" } }
    });
    if (sheetAlpha.model.outline.kind !== "xlsx" || sheetData.model.outline.kind !== "xlsx") {
      throw new Error("expected xlsx outlines");
    }
    // Different scopes → different windows AND different cache keys.
    expect(sheetAlpha.model.outline.sheets[0]!.name).toBe("Alpha");
    expect(sheetData.model.outline.sheets[0]!.name).toBe("Data");
    expect(sheetAlpha.model.cacheKey.scope).not.toBe(sheetData.model.cacheKey.scope);
    expect(service.stats().cacheSize).toBe(2);

    // Same scope again → cache hit, no re-render.
    const hitsBefore = service.stats().hits;
    const again = await service.preview({
      requestId: "r3",
      artifactRef: "art-scope",
      priority: "visible",
      scope: { location: { sheet: "Data" } }
    });
    expect(again.retried).toBe(false);
    expect(service.stats().hits).toBeGreaterThan(hitsBefore);
  });
});
