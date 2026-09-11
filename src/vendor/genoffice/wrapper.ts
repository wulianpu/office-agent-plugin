/**
 * Narrow GenOffice engine wrapper (§146 Adapter Boundary).
 *
 * The vendored engine sources (vendor/genoffice, Apache-2.0) are never
 * modified; this module exposes only the API surface the plugin depends on,
 * backed by prebuilt bundles from `npm run build:vendor` exposed as the
 * `genoffice-vendor` file: dependency. Sheets visual rendering (pptx-render)
 * and the Rust XLSX sidecar require a host/DOM or a cargo toolchain and are
 * intentionally outside this surface (§127 honest degradation).
 */

export { openPptx, savePptx } from "genoffice-vendor/pptx-engine.mjs";
export { parseDocx, saveDocx } from "genoffice-vendor/docx-engine.mjs";
export { buildRenderSlide } from "genoffice-vendor/pptx-render.mjs";
export type {
  OpenedPptx,
  PptxDeck,
  PptxSlide,
  PptxSlideElement,
  PptxTextBody,
  ParsedDocx,
  DocxBlock
} from "./genoffice-types.js";
export { elementText } from "./genoffice-types.js";

/** Capability probe: resolves true when the vendor bundles are loadable. */
export async function probeVendorEngines(): Promise<{ pptx: boolean; docx: boolean }> {
  const probe = async (specifier: string, symbol: string) => {
    try {
      const mod = (await import(specifier)) as Record<string, unknown>;
      return typeof mod[symbol] === "function";
    } catch {
      return false;
    }
  };
  return {
    pptx: await probe("genoffice-vendor/pptx-engine.mjs", "openPptx"),
    docx: await probe("genoffice-vendor/docx-engine.mjs", "parseDocx")
  };
}
