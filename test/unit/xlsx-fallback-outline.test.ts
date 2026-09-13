/**
 * Round 8 (issue #2): the XLSX JS fallback renderer must scope cell value
 * extraction to the cell's own XML. An empty cell (self-closed or empty
 * pair) previously inherited the NEXT cell's <v> — silently wrong preview
 * data whenever the Rust sidecar is unavailable.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { buildZip } from "../helpers/zip-builder.js";
import { renderXlsxOutline } from "../../src/preview/outline-renderers.js";

let dir: string;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "xlsx-fallback-test-"));
});

afterAll(async () => {
  await rm(dir, { recursive: true, force: true }).catch(() => undefined);
});

/** Minimal but structurally real workbook: workbook + rels + one sheet. */
async function writeWorkbook(
  name: string,
  sheetXml: string,
  options: { namespace?: boolean; sheetName?: string; sharedStrings?: string[] } = {}
): Promise<string> {
  const ns = options.namespace === false ? "" : "x:";
  const open = options.namespace === false ? "" : ` xmlns:x="main" xmlns:r="rel"`;
  const relOpen = options.namespace === false ? "" : ` xmlns:rel="relns"`;
  const sheetName = options.sheetName ?? "Sheet1";
  const entries: Array<{ name: string; data: string }> = [
    {
      name: "[Content_Types].xml",
      data: `<?xml version="1.0"?><Types xmlns="ct"/>`
    },
    {
      name: "xl/workbook.xml",
      data: `<?xml version="1.0"?><w${open}><${ns}sheets><${ns}sheet name="${sheetName}" sheetId="1" r:id="rId1"/></${ns}sheets></w>`
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      data: `<?xml version="1.0"?><Relationships${relOpen}><Relationship Id="rId1" Type="sheet" Target="worksheets/sheet1.xml"/></Relationships>`
    },
    {
      name: "xl/worksheets/sheet1.xml",
      data: `<?xml version="1.0"?><${ns}worksheet xmlns:x="main"><${ns}sheetData>${sheetXml}</${ns}sheetData></${ns}worksheet>`
    }
  ];
  if (options.sharedStrings) {
    entries.push({
      name: "xl/sharedStrings.xml",
      data: `<?xml version="1.0"?><sst>${options.sharedStrings
        .map((text) => `<si><t>${text}</t></si>`)
        .join("")}</sst>`
    });
  }
  const zip = buildZip(entries);
  const path = join(dir, `${name}.xlsx`);
  await writeFile(path, zip);
  return path;
}

describe("XLSX JS fallback cell scoping (round 8, issue #2)", () => {
  it("blank self-closed cell does not inherit the next cell's value (namespaced)", async () => {
    const path = await writeWorkbook("blank-ns", `<x:row r="1"><x:c r="A1"/><x:c r="B1"><x:v>42</x:v></x:c></x:row>`);
    const outline = await renderXlsxOutline(path);
    if (outline.kind !== "xlsx") throw new Error("expected xlsx outline");
    expect(outline.sheets[0]!.window[0]).toEqual(["", "42"]);
  });

  it("blank self-closed cell does not inherit the next cell's value (no namespace)", async () => {
    const path = await writeWorkbook(
      "blank-bare",
      `<row r="1"><c r="A1"/><c r="B1"><v>42</v></c></row>`,
      { namespace: false }
    );
    const outline = await renderXlsxOutline(path);
    if (outline.kind !== "xlsx") throw new Error("expected xlsx outline");
    expect(outline.sheets[0]!.window[0]).toEqual(["", "42"]);
  });

  it("explicit empty cell pair + column gap: no value bleed, gaps padded", async () => {
    const path = await writeWorkbook(
      "gap-ns",
      `<x:row r="1"><x:c r="A1"></x:c><x:c r="C1"><x:v>x</x:v></x:c></x:row>`
    );
    const outline = await renderXlsxOutline(path);
    if (outline.kind !== "xlsx") throw new Error("expected xlsx outline");
    expect(outline.sheets[0]!.window[0]).toEqual(["", "", "x"]);
  });

  it("shared strings resolve per cell, empties stay empty", async () => {
    const path = await writeWorkbook(
      "shared",
      `<row r="1"><c r="A1"/><c r="B1" t="s"><v>0</v></c></row>`,
      { namespace: false, sharedStrings: ["Hello"] }
    );
    const outline = await renderXlsxOutline(path);
    if (outline.kind !== "xlsx") throw new Error("expected xlsx outline");
    expect(outline.sheets[0]!.window[0]).toEqual(["", "Hello"]);
  });
});
