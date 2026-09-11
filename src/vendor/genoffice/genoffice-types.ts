/**
 * Minimal structural types for the GenOffice engines (§146 narrow surface).
 * Structural — the engine's exact internal types are not imported so engine
 * evolution inside the vendored tree cannot break compilation here.
 */

export interface PptxTextRun {
  text?: string;
}

export interface PptxTextBody {
  paragraphs?: Array<{ runs?: PptxTextRun[] }>;
}

export interface PptxSlideElement {
  type?: string;
  name?: string;
  /** Engine model: either a plain string or the structured text body. */
  text?: string | PptxTextBody;
  placeholder?: { type?: string } | null;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
}

/** Extract a flat string from the engine's text body shape. */
export function elementText(element: PptxSlideElement): string {
  if (typeof element.text === "string") return element.text;
  if (!element.text?.paragraphs) return "";
  return element.text.paragraphs
    .map((p) => (p.runs ?? []).map((r) => r.text ?? "").join(""))
    .join("\n");
}

export interface PptxSlide {
  elements: PptxSlideElement[];
  xml?: string;
}

export interface PptxDeck {
  slides: PptxSlide[];
  size: { cx: number; cy: number };
  originalHash?: string;
}

export interface OpenedPptx {
  deck: PptxDeck;
  archive: unknown;
}

export interface DocxBlock {
  type?: string;
  style?: string;
  text?: string;
  children?: DocxBlock[];
  [key: string]: unknown;
}

export interface ParsedDocx {
  blocks: DocxBlock[];
  [key: string]: unknown;
}
