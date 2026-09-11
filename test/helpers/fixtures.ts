/**
 * Fixture helpers: hermetic OOXML files built with the zip-builder, plus
 * workspace/plugin scaffolding shared by unit and integration suites.
 */

import { mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { buildZip } from "./zip-builder.js";
import { OfficePlugin } from "../../src/plugin/office-plugin.js";

export interface Workspace {
  root: string;
  /** Reassignable: crash-recovery tests reopen the plugin in-place. */
  plugin: OfficePlugin;
  cleanup: () => Promise<void>;
}

export async function openWorkspace(): Promise<Workspace> {
  const root = await mkdtemp(join(tmpdir(), "office-plugin-test-"));
  const plugin = await OfficePlugin.create({ workspaceRoot: join(root, "runtime"), engineDisabled: false, skipHostProbe: true });
  return {
    root,
    plugin,
    cleanup: async () => {
      await plugin.dispose();
      await rm(root, { recursive: true, force: true });
    }
  };
}

export async function writeDocxFixture(path: string, paragraphs: string[]): Promise<string> {
  const body = paragraphs
    .map(
      (text, i) =>
        `<w:p><w:pPr><w:pStyle w:val="Heading${i === 0 ? 1 : 2}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
    )
    .join("");
  const zip = buildZip([
    {
      name: "[Content_Types].xml",
      data: `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`
    },
    {
      name: "word/document.xml",
      data: `<?xml version="1.0"?><w:document xmlns:w="main"><w:body>${body}</w:body></w:document>`
    }
  ]);
  await writeFile(path, zip);
  return path;
}

export async function writeXlsxFixture(path: string, sheets: Array<{ name: string; rows: string[][] }>): Promise<string> {
  const sheetXml = (rows: string[][]) =>
    `<?xml version="1.0"?><worksheet xmlns="main"><sheetData>${rows
      .map(
        (row, r) =>
          `<row r="${r + 1}">${row
            .map((v, c) => `<c r="${colName(c)}${r + 1}"><v>${v}</v></c>`)
            .join("")}</row>`
      )
      .join("")}</sheetData></worksheet>`;
  const entries = [
    {
      name: "[Content_Types].xml",
      data: `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`
    },
    {
      name: "xl/workbook.xml",
      data: `<?xml version="1.0"?><workbook xmlns:r="rel"><sheets>${sheets
        .map((s, i) => `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
        .join("")}</sheets></workbook>`
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      data: `<?xml version="1.0"?><Relationships xmlns="rel">${sheets
        .map((_, i) => `<Relationship Id="rId${i + 1}" Target="worksheets/sheet${i + 1}.xml"/>`)
        .join("")}</Relationships>`
    },
    ...sheets.map((s, i) => ({ name: `xl/worksheets/sheet${i + 1}.xml`, data: sheetXml(s.rows) }))
  ];
  await writeFile(path, buildZip(entries));
  return path;
}

export async function writePptxFixture(
  path: string,
  slides: Array<Array<{ name: string; text: string }>>
): Promise<string> {
  const slideXml = (shapes: Array<{ name: string; text: string }>) =>
    `<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:spTree>${shapes
      .map(
        (s) =>
          `<p:sp><p:nvSpPr><p:cNvPr id="1" name="${s.name}"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>${s.text}</a:t></a:r></a:p></p:txBody></p:sp>`
      )
      .join("")}</p:spTree></p:sld>`;
  const entries = [
    {
      name: "[Content_Types].xml",
      data: `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`
    },
    {
      name: "ppt/presentation.xml",
      data: `<?xml version="1.0"?><p:presentation xmlns:p="p"><p:sldIdLst>${slides
        .map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`)
        .join("")}</p:sldIdLst></p:presentation>`
    },
    ...slides.map((shapes, i) => ({ name: `ppt/slides/slide${i + 1}.xml`, data: slideXml(shapes) }))
  ];
  await writeFile(path, buildZip(entries));
  return path;
}

export function colName(index: number): string {
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}
