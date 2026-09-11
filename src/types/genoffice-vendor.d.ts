/**
 * Ambient declarations for the prebuilt GenOffice engine bundles.
 */

declare module "genoffice-vendor/pptx-engine.mjs" {
  export function openPptx(bytes: Uint8Array): Promise<import("./vendor/genoffice/genoffice-types.js").OpenedPptx>;
  export function savePptx(opened: import("./vendor/genoffice/genoffice-types.js").OpenedPptx): Promise<Uint8Array>;
}

declare module "genoffice-vendor/docx-engine.mjs" {
  export function parseDocx(bytes: Uint8Array): Promise<import("./vendor/genoffice/genoffice-types.js").ParsedDocx>;
  export function saveDocx(
    parsed: import("./vendor/genoffice/genoffice-types.js").ParsedDocx,
    options?: Record<string, unknown>
  ): Promise<Uint8Array>;
}

declare module "genoffice-vendor/pptx-render.mjs" {
  import type { PptxDeck } from "./vendor/genoffice/genoffice-types.js";
  // The render layer's structural output — opaque to us beyond what the SVG
  // adapter consumes, hence a loose record type.
  export function buildRenderSlide(
    slide: unknown,
    size: PptxDeck["size"],
    options: { fitWidthPx: number; slideNo?: number }
  ): unknown;
}
