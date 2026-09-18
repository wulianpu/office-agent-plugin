/**
 * Golden Corpus generator (§133): small/medium files per format via officecli
 * plus a large stored-ZIP XLSX built locally (fast generation, big parse).
 *
 * Usage: node tools/corpus/generate.mjs [targetDir]
 * Output: .corpus/ with manifest.json
 */

import { mkdir, writeFile } from "node:fs/promises";
import { join, resolve } from "node:path";
import { execFile } from "node:child_process";
import { pathToFileURL } from "node:url";
import { promisify } from "node:util";
import { crc32 } from "node:zlib";

const run = promisify(execFile);

/** Stored-entry ZIP builder (no compression; fast for large generated parts). */
function buildZip(entries) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const entry of entries) {
    const data = typeof entry.data === "string" ? Buffer.from(entry.data, "utf8") : entry.data;
    const nameBuf = Buffer.from(entry.name, "utf8");
    const crc = crc32(data) >>> 0;
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(0, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0x21, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    locals.push(local, nameBuf, data);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0, 8);
    central.writeUInt16LE(0, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0x21, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, nameBuf);
    offset += 30 + nameBuf.length + data.length;
  }
  const cdBuf = Buffer.concat(centrals);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(cdBuf.length, 12);
  eocd.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, cdBuf, eocd]);
}

// Windows: the npm global install only provides a .cmd shim — bare execFile
// cannot spawn it and shell mode mangles the JSON argv (CI produced 1 of 6
// corpus files this way). Prefer the dist adapter, which resolves the shim
// to its JS entry; fall back to bare execFile where dist is absent.
let adapter;
try {
  adapter = new (
    await import(pathToFileURL(join(process.cwd(), "dist", "agent", "officecli", "officecli-adapter.js")).href)
  ).OfficeCliAdapter({ timeoutMs: 300_000 });
} catch {
  adapter = null;
}

async function officecli(args, cwd) {
  if (adapter) {
    return JSON.stringify(await adapter.run(args));
  }
  const { stdout } = await run("officecli", args, { cwd, windowsHide: true, maxBuffer: 64 * 1024 * 1024 });
  return stdout;
}

async function createViaOfficecli(dir, name, items) {
  const file = join(dir, name);
  await officecli(["create", file, "--json"]).catch(() => undefined);
  if (items.length > 0) {
    await officecli(["batch", file, "--commands", JSON.stringify(items), "--json"]).catch(() => undefined);
  }
  await officecli(["close", file, "--json"]).catch(() => undefined);
  return name;
}

function largeSheetXml(rows, cols) {
  const parts = ['<?xml version="1.0"?><worksheet xmlns="main"><sheetData>'];
  for (let r = 1; r <= rows; r++) {
    let row = `<row r="${r}">`;
    for (let c = 1; c <= cols; c++) {
      row += `<c r="${colName(c)}${r}"><v>${(r * c) % 9973}</v></c>`;
    }
    row += "</row>";
    parts.push(row);
  }
  parts.push("</sheetData></worksheet>");
  return parts.join("");
}

function colName(index) {
  let n = index;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export async function generateCorpus(targetDir = ".corpus") {
  const dir = resolve(targetDir);
  await mkdir(dir, { recursive: true });
  const manifest = { generatedAt: new Date().toISOString(), files: [] };

  let officecliAvailable = true;
  try {
    await officecli(["--version"]);
  } catch {
    officecliAvailable = false;
  }
  manifest.officecli = officecliAvailable;

  if (officecliAvailable) {
    manifest.files.push(await createViaOfficecli(dir, "docx-small.docx", [
      { command: "add", parent: "/body", type: "paragraph", props: { text: "Small doc body" } }
    ]));
    const mediumDocx = [];
    for (let i = 0; i < 120; i++) {
      mediumDocx.push({ command: "add", parent: "/body", type: "paragraph", props: { text: `Paragraph ${i} with some content to give the document realistic size.` } });
    }
    manifest.files.push(await createViaOfficecli(dir, "docx-medium.docx", mediumDocx));

    manifest.files.push(await createViaOfficecli(dir, "pptx-small.pptx", [
      { command: "add", parent: "/", type: "slide" },
      { command: "add", parent: "/slide[1]", type: "shape", props: { text: "Title" } }
    ]));
    const mediumPptx = [];
    for (let s = 1; s <= 25; s++) {
      mediumPptx.push({ command: "add", parent: "/", type: "slide" });
      mediumPptx.push({ command: "add", parent: `/slide[${s}]`, type: "shape", props: { text: `Slide ${s} body text` } });
      mediumPptx.push({ command: "add", parent: `/slide[${s}]`, type: "shape", props: { text: `Slide ${s} footer` } });
    }
    manifest.files.push(await createViaOfficecli(dir, "pptx-medium.pptx", mediumPptx));

    manifest.files.push(await createViaOfficecli(dir, "xlsx-small.xlsx", [
      { command: "add", parent: "/", type: "sheet", props: { name: "S1" } },
      { command: "set", path: "/sheet[1]/cell[A1]", props: { value: "a" } }
    ]));
  }

  // Large XLSX (§138): rows × 26 cols stored locally — tens of MB, fast.
  // The engine rejects opens above 3,000,000 XML elements (1.0.151 memory
  // guard; empirically tripped by the CI resident on larger files). 45k rows
  // ≈1.2M cells / ~34MB — safely under the guard on every counting basis,
  // still a tens-of-MB stored workbook. CORPUS_LARGE_ROWS overrides.
  const rows = Number(process.env.CORPUS_LARGE_ROWS ?? 45_000);
  const large = buildZip([
    { name: "[Content_Types].xml", data: '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>' },
    { name: "xl/workbook.xml", data: '<?xml version="1.0"?><workbook xmlns:r="rel"><sheets><sheet name="Big" sheetId="1" r:id="rId1"/></sheets></workbook>' },
    { name: "xl/_rels/workbook.xml.rels", data: '<?xml version="1.0"?><Relationships xmlns="rel"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>' },
    { name: "xl/worksheets/sheet1.xml", data: largeSheetXml(rows, 26) }
  ]);
  await writeFile(join(dir, "xlsx-large.xlsx"), large);
  manifest.files.push("xlsx-large.xlsx");
  manifest.largeXlsxRows = rows;
  manifest.largeXlsxBytes = large.length;

  await writeFile(join(dir, "manifest.json"), JSON.stringify(manifest, null, 2));
  return manifest;
}

// CLI entry
if (process.argv[1] && resolve(process.argv[1]) === resolve(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, "$1"))) {
  generateCorpus(process.argv[2] ?? ".corpus")
    .then((m) => {
      console.log(`corpus generated: ${m.files.length} files (${m.largeXlsxBytes ?? 0} large bytes)`);
    })
    .catch((error) => {
      console.error(error);
      process.exit(1);
    });
}
