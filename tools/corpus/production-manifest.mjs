/**
 * Issue #15 producer/format matrix: promote suite-generated corpus files from
 * the staging dir (tools/corpus/new-origin-corpus.ps1) into
 * corpus/production/ and update the versioned manifest.
 *
 * The manifest stays the fail-closed evidence contract enforced by
 * test/integration/production-compat.test.ts — this tool only ever adds or
 * refreshes entries for staged files; existing entries are never removed.
 *
 * Usage:
 *   node tools/corpus/production-manifest.mjs <stagingDir> <producer> <origin>
 *   node tools/corpus/production-manifest.mjs .corpus/origin-staging wps wps-suite-synthetic
 */

import { createHash } from "node:crypto";
import { existsSync } from "node:fs";
import { copyFile, mkdir, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { extname, join } from "node:path";

const [stagingDir, producer, origin] = process.argv.slice(2);
if (!stagingDir || !producer || !origin) {
  console.error("usage: node tools/corpus/production-manifest.mjs <stagingDir> <producer> <origin>");
  process.exit(1);
}
if (!["office", "wps", "third-party"].includes(producer)) {
  console.error(`illegal producer: ${producer}`);
  process.exit(1);
}

const PRODUCER_DIR = join(process.cwd(), "corpus", "production");
const MANIFEST_PATH = join(PRODUCER_DIR, "manifest.json");

const featuresFor = (format) => {
  switch (format) {
    case ".docx":
      return ["styles", "headings", "tables"];
    case ".xlsx":
      return ["formulas", "multiple-sheets"];
    case ".pptx":
      return ["slide-master", "placeholders", "multiple-slides"];
    default:
      throw new Error(`no feature mapping for format ${format} — fail-closed`);
  }
};

if (!existsSync(stagingDir)) {
  console.error(`staging dir not found: ${stagingDir}`);
  process.exit(1);
}
const staged = (await readdir(stagingDir)).filter((f) => /^\w+-\w+\.(docx|xlsx|pptx)$/.test(f));
if (staged.length === 0) {
  console.error(`no generated corpus files in ${stagingDir}`);
  process.exit(1);
}

await mkdir(PRODUCER_DIR, { recursive: true });
const manifest = JSON.parse(await readFile(MANIFEST_PATH, "utf8"));

for (const file of staged) {
  const src = join(stagingDir, file);
  const dest = join(PRODUCER_DIR, file);
  await copyFile(src, dest);
  const sha256 = createHash("sha256").update(await readFile(dest)).digest("hex");
  const entry = {
    file,
    producer,
    origin,
    features: featuresFor(extname(file)),
    sha256,
    expect: "full"
  };
  const index = manifest.files.findIndex((f) => f.file === file);
  if (index >= 0) {
    manifest.files[index] = entry;
  } else {
    manifest.files.push(entry);
  }
  console.log(`promoted: ${file} (sha256 ${sha256.slice(0, 12)}…)`);
}

manifest.version = 1;
await writeFile(MANIFEST_PATH, JSON.stringify(manifest, null, 2) + "\n");
await rm(stagingDir, { recursive: true, force: true }).catch(() => undefined);
console.log(`manifest updated: ${MANIFEST_PATH} (${manifest.files.length} entries)`);
