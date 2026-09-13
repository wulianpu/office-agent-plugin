/**
 * Headless preview renderers for the three formats (§28–§30). All extraction
 * is streaming and bounded — the renderer's working set stays proportional to
 * the visible window, never the file size (§31, PERF-13).
 */

import type { PreviewOutline, PreviewScope } from "../contracts/preview.js";
import { concatenatedTagTexts, decodeXmlEntities, extractTagTexts, splitRows, splitTagged } from "../support/xml-lite.js";
import { readZipEntry, readZipIndex, streamZipEntry } from "../artifact/scanner/zip.js";

const DOCX_MAX_BLOCKS = 400;
const XLSX_MAX_ROWS = 60;
const XLSX_MAX_COLS = 24;
const PPTX_MAX_SLIDES = 40;

export async function renderDocxOutline(path: string, scope?: PreviewScope): Promise<PreviewOutline> {
  const anchor = Math.max(0, Math.floor(scope?.location?.block ?? 0));
  const maxBlocks = Math.min(DOCX_MAX_BLOCKS, Math.max(1, scope?.maxEntries ?? DOCX_MAX_BLOCKS));
  const index = await readZipIndex(path);
  const blocks: Array<{ index: number; style?: string; text: string }> = [];
  const carry = { pending: "" };
  let carryText = "";
  let done = false;
  let seen = 0; // absolute paragraph counter (anchor windowing, round 10)

  for await (const chunk of streamZipEntry(path, index, "word/document.xml")) {
    const text = chunk.toString("utf8");
    // Paragraph-level split: <w:p ...>…</w:p>
    let data = carryText + text;
    carryText = "";
    let searchFrom = 0;
    for (;;) {
      if (blocks.length >= maxBlocks) {
        done = true;
        break;
      }
      const openIdx = data.indexOf("<w:p ", searchFrom) >= 0
        ? data.indexOf("<w:p ", searchFrom)
        : data.indexOf("<w:p>", searchFrom);
      if (openIdx < 0) {
        carryText = data.slice(Math.max(0, data.length - 8));
        break;
      }
      const closeIdx = data.indexOf("</w:p>", openIdx);
      if (closeIdx < 0) {
        carryText = data.slice(openIdx);
        break;
      }
      const para = data.slice(openIdx, closeIdx);
      const absolute = seen++;
      if (absolute >= anchor) {
        blocks.push({
          index: absolute,
          style: para.match(/<w:pStyle w:val="([^"]+)"/)?.[1],
          text: extractTagTexts(para, "w:t", { pending: "" }).join("")
        });
      }
      searchFrom = closeIdx + 6;
    }
    if (done) break;
  }
  return { kind: "docx", blocks };
}

export async function renderXlsxOutline(path: string, scope?: PreviewScope): Promise<PreviewOutline> {
  const index = await readZipIndex(path);
  const sheets: Array<{ name: string; rowCount: number; rowCountExact?: boolean; window: string[][] }> = [];

  const workbookXml = (
    await readZipEntry(path, index, "xl/workbook.xml", 16 * 1024 * 1024).catch(() => Buffer.alloc(0))
  ).toString("utf8");
  // Engine writers may namespace the tags (<x:sheet …>) and use arbitrary
  // relationship ids (not just rIdN) — match the tag, then the attributes.
  const sheetTags = [...workbookXml.matchAll(/<(?:[\w.-]+:)?sheet\b([^>]*?)\/?>/g)]
    .map((m) => {
      const attrs = m[1] ?? "";
      return {
        name: attrs.match(/\bname="([^"]+)"/)?.[1],
        rid: attrs.match(/\br:id="([^"]+)"/)?.[1]
      };
    })
    .filter((s): s is { name: string; rid: string } => Boolean(s.name && s.rid));

  // sharedStrings (bounded): the INDEX UNIT is <si>, not <t>. A Rich Text
  // <si> carries multiple <r><t> runs that concatenate into ONE entry —
  // per-<t> indexing shifted every later index after the first rich entry
  // (round 9, P1-high). Streaming <si> records keep the 100k bound without
  // materializing the workbook.
  const shared: string[] = [];
  if (index.entryByName.has("xl/sharedStrings.xml")) {
    const carry = { pending: "" };
    for await (const chunk of streamZipEntry(path, index, "xl/sharedStrings.xml")) {
      for (const record of splitTagged(chunk.toString("utf8"), "si", carry)) {
        if (shared.length >= 100_000) break;
        shared.push(concatenatedTagTexts(record, "t"));
      }
      if (shared.length >= 100_000) break;
    }
  }

  const workbookRels = (
    await readZipEntry(path, index, "xl/_rels/workbook.xml.rels", 8 * 1024 * 1024).catch(() => Buffer.alloc(0))
  ).toString("utf8");
  const relTargets = new Map(
    [...workbookRels.matchAll(/<(?:[\w.-]+:)?Relationship\b([^>]*?)\/?>/g)]
      .map((m) => {
        const attrs = m[1] ?? "";
        const id = attrs.match(/\bId="([^"]+)"/)?.[1];
        const target = attrs.match(/\bTarget="([^"]+)"/)?.[1];
        return id && target ? ([id, target.replace(/^\//, "")] as [string, string]) : undefined;
      })
      .filter((e): e is [string, string] => Boolean(e))
  );

  // Round 10 scoped windowing: an explicit sheet selects it (falling back to
  // the leading sheets when unresolvable — a degraded-but-honest default);
  // range/maxEntries bound the window; maxEntries never lifts the hard caps.
  const wantedSheet = scope?.location?.sheet;
  const range = scope?.location?.range;
  const maxRows = Math.min(
    XLSX_MAX_ROWS,
    Math.max(1, scope?.maxEntries ?? (range ? range.toRow - range.fromRow + 1 : XLSX_MAX_ROWS))
  );
  const maxCols = range ? Math.min(XLSX_MAX_COLS, range.toCol - range.fromCol + 1) : XLSX_MAX_COLS;
  const fromRow = range ? Math.max(1, range.fromRow) : 1;
  const fromCol = range ? Math.max(1, range.fromCol) : 1;
  // An explicit range is a HARD row boundary: maxEntries may narrow it but
  // never extend the window past toRow.
  const toRow = range ? Math.max(fromRow, range.toRow) : Number.POSITIVE_INFINITY;

  // Round 10 reopen (fail-closed): an EXPLICIT sheet that does not exist
  // yields an EMPTY scoped result — never the workbook's real sheets under
  // a cache key naming the missing sheet.
  if (wantedSheet && !sheetTags.some((t) => t.name === wantedSheet)) {
    return { kind: "xlsx", sheets: [] };
  }
  const ordered = wantedSheet
    ? [
        sheetTags.find((t) => t.name === wantedSheet)!,
        ...sheetTags.filter((t) => t.name !== wantedSheet)
      ]
    : sheetTags;

  let sheetIdx = 0;
  for (const tag of ordered) {
    if (sheetIdx >= 4) break; // preview window covers the first sheets
    const name = tag.name;
    const target = relTargets.get(tag.rid);
    const part = target
      ? target.startsWith("xl/")
        ? target
        : `xl/${target}`
      : `xl/worksheets/sheet${sheetIdx + 1}.xml`;
    const window: string[][] = [];
    let rowCount = 0;
    // Round 10 rowCount honesty: a declared <dimension ref="A1:C12345"/> is
    // the EXACT extent; without it the scanned probe count is only a lower
    // bound (rowCountExact: false) — never presented as a precise total.
    let declaredRows: number | undefined;
    if (index.entryByName.has(part)) {
      const carry = { pending: "" };
      let dimensionScan = "";
      let hardStop = false;
      for await (const chunk of streamZipEntry(path, index, part)) {
        if (declaredRows === undefined && dimensionScan.length < 16 * 1024) {
          dimensionScan += chunk.toString("utf8");
          const ref = dimensionScan.match(/<(?:[\w.-]+:)?dimension\b[^>]*ref="([A-Z]+\d+):([A-Z]+)(\d+)"/);
          if (ref) declaredRows = Number(ref[3]);
        }
        for (const row of splitRows(chunk.toString("utf8"), carry)) {
          rowCount++;
          const rowNumber = row.match(/<(?:[\w.-]+:)?row\b[^>]*\br="(\d+)"/)?.[1];
          const absolute = rowNumber ? Number(rowNumber) : rowCount;
          if (absolute > toRow) {
            hardStop = true; // explicit range boundary reached
            break;
          }
          if (absolute >= fromRow && window.length < maxRows) {
            window.push(parseRowCells(row, shared, fromCol, maxCols));
          }
          if (rowCount > 500_000) break; // row-count probe cap
        }
        if (hardStop) break;
        if (window.length >= maxRows && rowCount > maxRows * 4) break;
      }
    }
    const exact = declaredRows !== undefined;
    const rowCountOut = exact ? Math.max(declaredRows!, rowCount) : rowCount;
    sheets.push({ name, rowCount: rowCountOut, rowCountExact: exact, window });
    sheetIdx++;
  }
  return { kind: "xlsx", sheets };
}

function parseRowCells(rowXml: string, shared: string[], fromCol = 1, maxCols = XLSX_MAX_COLS): string[] {
  // Round 10 reopen: placement is by ABSOLUTE column ref directly into a
  // viewport-local slot (local = absoluteCol - (fromCol-1)). No padding from
  // column A to the window origin and no absolute column ceiling — a range
  // starting at IW works with a working set proportional to the WINDOW, not
  // to fromCol.
  const cells: string[] = [];
  let unrefSlot = 0;
  // Namespace-tolerant cell matchers: engine sheets use <x:c>/<x:v>/<x:is>.
  // The open-tag alternative requires a non-`/` before `>` so a self-closed
  // cell can never swallow the NEXT cell's content up to its close tag —
  // that misparse chained the follower's value into blank cells (round 8).
  const cellRe =
    /<(?:[\w.-]+:)?c ([^>]*[^/>])>[\s\S]*?<\/(?:[\w.-]+:)?c>|<(?:[\w.-]+:)?c ([^>]*?)\/>/g;
  for (const m of rowXml.matchAll(cellRe)) {
    const cellXml = m[0];
    const attrs = (m[1] ?? m[2] ?? "") + ">";
    const refMatch = attrs.match(/r="([A-Z]+)\d+"/);
    const ref = refMatch?.[1] ?? "";
    // Value extraction is scoped to THIS cell's XML only — slicing from
    // m.index across the row let an empty cell inherit the next cell's <v>.
    let value = "";
    const vMatch = cellXml.match(/<(?:[\w.-]+:)?v>([^<]*)<\/(?:[\w.-]+:)?v>/);
    if (vMatch) {
      const raw = vMatch[1]!;
      const isShared = /t="s"/.test(attrs);
      // shared[] entries are already entity-decoded (concatenatedTagTexts);
      // decode only the raw <v> literal here.
      value = isShared ? (shared[Number(raw)] ?? "") : decodeXmlEntities(raw);
    } else {
      // Inline string: may itself be Rich Text (<is><r><t>…</t></r>…</is>) —
      // concatenate every run inside the <is>, entity-decoded (round 9).
      const inline = cellXml.match(/<(?:[\w.-]+:)?is>[\s\S]*?<\/(?:[\w.-]+:)?is>/);
      if (inline) value = concatenatedTagTexts(inline[0], "t");
    }
    const absoluteCol = ref ? columnToIndex(ref) : fromCol - 1 + unrefSlot++;
    const local = absoluteCol - (fromCol - 1);
    if (local < 0 || local >= maxCols) continue; // outside the window: skip, never pad
    cells[local] = value;
  }
  for (let i = 0; i < cells.length; i++) if (cells[i] === undefined) cells[i] = "";
  return cells;
}

function columnToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export async function renderPptxOutline(path: string, scope?: PreviewScope): Promise<PreviewOutline> {
  // Round 10 scoped windowing: the slide anchor opens the window at the
  // requested 1-based slide; maxEntries caps it. The hard ceiling stays.
  // Round 10 reopen: an anchor beyond the deck clamps to the LAST slide —
  // the same policy as the engine outline/SVG window (svgWindowOf), never
  // an empty outline in one view and a last-page window in another.
  const anchor = Math.max(1, Math.floor(scope?.location?.slide ?? 1));
  const maxSlides = Math.min(PPTX_MAX_SLIDES, Math.max(1, scope?.maxEntries ?? PPTX_MAX_SLIDES));
  const index = await readZipIndex(path);
  const all = index.entries
    .filter((e) => /^ppt\/slides\/slide\d+\.xml$/.test(e.name))
    .sort((a, b) => slideNo(a.name) - slideNo(b.name));
  const inWindow = all.filter((e) => slideNo(e.name) >= anchor && slideNo(e.name) < anchor + maxSlides);
  const slideEntries = inWindow.length > 0 ? inWindow : all.slice(-1);

  const slides: Array<{ index: number; shapes: Array<{ name?: string; text?: string }> }> = [];
  for (const entry of slideEntries) {
    const xml = (await readZipEntry(path, index, entry.name, 32 * 1024 * 1024).catch(() => Buffer.alloc(0))).toString("utf8");
    const texts = extractTagTexts(xml, "a:t", { pending: "" });
    const shapes: Array<{ name?: string; text?: string }> = [];
    for (const m of xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
      const shapeXml = m[0];
      const nameMatch = shapeXml.match(/<p:cNvPr id="\d+" name="([^"]*)"/);
      const shapeTexts = extractTagTexts(shapeXml, "a:t", { pending: "" });
      shapes.push({
        name: nameMatch?.[1],
        text: shapeTexts.join(" ") || undefined
      });
    }
    const slideIndex = slideNo(entry.name);
    slides.push({ index: slideIndex, shapes: shapes.length > 0 ? shapes : [{ text: texts.join(" ") || undefined }] });
  }
  return { kind: "pptx", slides };
}

function slideNo(name: string): number {
  return Number(name.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
}
