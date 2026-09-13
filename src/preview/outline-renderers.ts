/**
 * Headless preview renderers for the three formats (§28–§30). All extraction
 * is streaming and bounded — the renderer's working set stays proportional to
 * the visible window, never the file size (§31, PERF-13).
 */

import type { PreviewOutline } from "../contracts/preview.js";
import { extractTagTexts, splitRows } from "../support/xml-lite.js";
import { readZipEntry, readZipIndex, streamZipEntry } from "../artifact/scanner/zip.js";

const DOCX_MAX_BLOCKS = 400;
const XLSX_MAX_ROWS = 60;
const XLSX_MAX_COLS = 24;
const PPTX_MAX_SLIDES = 40;

export async function renderDocxOutline(path: string): Promise<PreviewOutline> {
  const index = await readZipIndex(path);
  const blocks: Array<{ index: number; style?: string; text: string }> = [];
  const carry = { pending: "" };
  let carryText = "";
  let done = false;

  for await (const chunk of streamZipEntry(path, index, "word/document.xml")) {
    const text = chunk.toString("utf8");
    // Paragraph-level split: <w:p ...>…</w:p>
    let data = carryText + text;
    carryText = "";
    let searchFrom = 0;
    for (;;) {
      if (blocks.length >= DOCX_MAX_BLOCKS) {
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
      const texts = extractTagTexts(para, "w:t", { pending: "" });
      const joined = texts.join("");
      const styleMatch = para.match(/<w:pStyle w:val="([^"]+)"/);
      blocks.push({
        index: blocks.length,
        style: styleMatch?.[1],
        text: joined
      });
      searchFrom = closeIdx + 6;
    }
    if (done) break;
  }
  return { kind: "docx", blocks };
}

export async function renderXlsxOutline(path: string): Promise<PreviewOutline> {
  const index = await readZipIndex(path);
  const sheets: Array<{ name: string; rowCount: number; window: string[][] }> = [];

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

  // sharedStrings (bounded): map index → text.
  const shared: string[] = [];
  if (index.entryByName.has("xl/sharedStrings.xml")) {
    const carry = { pending: "" };
    for await (const chunk of streamZipEntry(path, index, "xl/sharedStrings.xml")) {
      for (const text of extractTagTexts(chunk.toString("utf8"), "t", carry)) {
        if (shared.length < 100_000) shared.push(text);
      }
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

  let sheetIdx = 0;
  for (const tag of sheetTags) {
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
    if (index.entryByName.has(part)) {
      const carry = { pending: "" };
      for await (const chunk of streamZipEntry(path, index, part)) {
        for (const row of splitRows(chunk.toString("utf8"), carry)) {
          rowCount++;
          if (window.length < XLSX_MAX_ROWS) {
            window.push(parseRowCells(row, shared));
          }
          if (rowCount > 500_000) break; // row-count probe cap
        }
        if (window.length >= XLSX_MAX_ROWS && rowCount > XLSX_MAX_ROWS * 4) break;
      }
    }
    sheets.push({ name, rowCount, window });
    sheetIdx++;
  }
  return { kind: "xlsx", sheets };
}

function parseRowCells(rowXml: string, shared: string[]): string[] {
  const cells: string[] = [];
  // Namespace-tolerant cell matchers: engine sheets use <x:c>/<x:v>/<x:is>.
  const cellRe = /<(?:[\w.-]+:)?c ([^>]*)>[\s\S]*?<\/(?:[\w.-]+:)?c>|<(?:[\w.-]+:)?c ([^>]*)\/>/g;
  for (const m of rowXml.matchAll(cellRe)) {
    const attrs = (m[1] ?? m[2] ?? "") + ">";
    const refMatch = attrs.match(/r="([A-Z]+)\d+"/);
    const ref = refMatch?.[1] ?? "";
    let value = "";
    const rest = rowXml.slice(m.index ?? 0);
    const vMatch = rest.match(/<(?:[\w.-]+:)?v>([^<]*)<\/(?:[\w.-]+:)?v>/);
    if (vMatch) {
      const raw = vMatch[1]!;
      const isShared = /t="s"/.test(attrs);
      value = isShared ? (shared[Number(raw)] ?? "") : raw;
    } else {
      const inline = rest.match(/<(?:[\w.-]+:)?is>\s*<(?:[\w.-]+:)?t[^>]*>([^<]*)<\/(?:[\w.-]+:)?t>/);
      if (inline) value = inline[1]!;
    }
    const colIndex = ref ? columnToIndex(ref) : cells.length;
    while (cells.length < colIndex && cells.length < XLSX_MAX_COLS) cells.push("");
    if (cells.length < XLSX_MAX_COLS) cells.push(value);
    if (cells.length >= XLSX_MAX_COLS) break;
  }
  return cells;
}

function columnToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export async function renderPptxOutline(path: string): Promise<PreviewOutline> {
  const index = await readZipIndex(path);
  const slideEntries = index.entries
    .filter((e) => /^ppt\/slides\/slide\d+\.xml$/.test(e.name))
    .sort((a, b) => slideNo(a.name) - slideNo(b.name))
    .slice(0, PPTX_MAX_SLIDES);

  const slides: Array<{ index: number; shapes: Array<{ name?: string; text?: string }> }> = [];
  let i = 0;
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
    slides.push({ index: i + 1, shapes: shapes.length > 0 ? shapes : [{ text: texts.join(" ") || undefined }] });
    i++;
  }
  return { kind: "pptx", slides };
}

function slideNo(name: string): number {
  return Number(name.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
}
