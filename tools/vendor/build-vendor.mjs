/**
 * Vendor build (§146 GenOffice Fork Policy): bundle the vendored GenOffice
 * engine packages (unmodified sources) into single-file ESM modules that our
 * runtime imports through the narrow wrapper in src/vendor/genoffice/.
 *
 * Usage: node tools/vendor/build-vendor.mjs
 */

import { build } from "esbuild";
import { mkdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const root = join(dirname(fileURLToPath(import.meta.url)), "..", "..");
const vendor = join(root, "vendor", "genoffice", "packages");
const outDir = join(root, "vendor-bundle");

const EXTERNAL = ["jszip", "fast-xml-parser", "utif2"];

await mkdir(outDir, { recursive: true });

/** @genoffice/pptx-engine subpath → vendored TS source (workspace exports). */
function pptxEngineAliases() {
  const engineDir = join(vendor, "pptx-engine", "src");
  const aliases = {
    "@genoffice/pptx-engine": join(engineDir, "index.ts"),
  };
  for (const sub of ["custgeom", "table-grid", "identity", "background-promote"]) {
    aliases[`@genoffice/pptx-engine/${sub}`] = join(engineDir, `${sub}.ts`);
  }
  return aliases;
}

const targets = [
  {
    entry: join(vendor, "pptx-engine", "src", "index.ts"),
    outfile: join(outDir, "pptx-engine.mjs"),
    alias: {},
  },
  {
    entry: join(vendor, "pptx-render", "src", "index.ts"),
    outfile: join(outDir, "pptx-render.mjs"),
    alias: pptxEngineAliases(),
  },
  {
    entry: join(vendor, "docx-engine", "src", "index.ts"),
    outfile: join(outDir, "docx-engine.mjs"),
    alias: pptxEngineAliases(),
  },
];

for (const target of targets) {
  const result = await build({
    entryPoints: [target.entry],
    bundle: true,
    format: "esm",
    platform: "node",
    target: "node22",
    outfile: target.outfile,
    external: EXTERNAL,
    alias: target.alias,
    sourcemap: true,
    logLevel: "warning",
  });
  const errors = result.errors ?? [];
  if (errors.length > 0) {
    throw new Error(`vendor build failed for ${target.outfile}: ${errors.length} error(s)`);
  }
  console.log(`bundled ${target.outfile.replace(root, ".")}`);
}
