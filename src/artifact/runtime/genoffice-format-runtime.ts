/**
 * GenOffice-backed FormatRuntime (§26, §146): parse artifacts through the
 * vendored GenOffice engines (unmodified, bundled) so ArtifactContexts carry
 * the real engine data model (deck / blocks) instead of a bare ZIP index.
 *
 * §146 compliance: engine cores are never edited; this file is exactly the
 * "ArtifactContext integration / EditorAdapter" layer the design adds. The
 * basic runtime remains the fallback when the vendor bundles are absent.
 */

import type {
  ArtifactBuildInput,
  ArtifactContext,
  FormatRuntime,
  MemoryTrimLevel
} from "../../contracts/artifact.js";
import { fileFingerprint } from "../../support/fsx.js";
import { BasicFormatRuntime } from "./basic-format-runtime.js";
import { openPptx, parseDocx } from "../../vendor/genoffice/wrapper.js";
import type { DocxBlock, ParsedDocx, PptxDeck } from "../../vendor/genoffice/wrapper.js";

export const GENOFFICE_RENDERER_VERSION = "genoffice-1";

/**
 * Parsed models ride INSIDE the context enrichment (§25): identity wrappers
 * can be rebound on promotion while the underlying engine model stays shared
 * — no WeakMap identity trap (P1-high-A).
 */
const DECK_KEY = "genoffice:deck";
const DOC_KEY = "genoffice:doc";

export class GenOfficePptxFormatRuntime implements FormatRuntime {
  readonly format = "pptx" as const;
  /** §127: degraded fallback when the engine cannot make sense of the bytes. */
  private readonly fallback = new BasicFormatRuntime("pptx");

  async initialize(): Promise<void> {}

  async createArtifactContext(input: ArtifactBuildInput): Promise<ArtifactContext> {
    try {
      return await this.createFromEngine(input);
    } catch {
      await this.fallback.initialize();
      return this.fallback.createArtifactContext(input);
    }
  }

  private async createFromEngine(input: ArtifactBuildInput): Promise<ArtifactContext> {
    const path = this.pathResolver(input.artifactRef);
    const fingerprint = await fileFingerprint(path);
    if (input.signal?.aborted) throw new DOMException("aborted", "AbortError");
    // Zero-copy: wrap the read buffer's ArrayBuffer (no duplicate payload in memory).
    const buffer = await this.readFile(path);
    const bytes = new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
    const opened = await openPptx(bytes);
    if (opened.deck.slides.length === 0) {
      // Engine parsed but found nothing (e.g. minimal/nonstandard packages):
      // degrade to the basic context so preview falls back to zip streaming.
      throw new Error("engine produced an empty deck");
    }
    if (input.signal?.aborted) throw new DOMException("aborted", "AbortError");

    const context: ArtifactContext = {
      artifactRef: input.artifactRef,
      version: { artifactRef: input.artifactRef, fingerprint },
      format: "pptx",
      consistency: input.consistency,
      profile: "full",
      rendererVersion: GENOFFICE_RENDERER_VERSION,
      lastAccessAt: Date.now(),
      // §101 honesty: deck elements are the resident cost, not a flat 8KB.
      estimatedResidentBytes: estimateDeckBytes(opened.deck, buffer.byteLength),
      enrichment: new Map<string, unknown>([
        ["engine", "genoffice"],
        ["slideCount", opened.deck.slides.length],
        ["deckSize", opened.deck.size],
        [DECK_KEY, opened.deck]
      ])
    };
    return context;
  }

  async trimMemory(level: MemoryTrimLevel): Promise<void> {
    await this.fallback.trimMemory(level);
  }

  async dispose(): Promise<void> {
    await this.fallback.dispose();
  }

  private pathResolver: (ref: string) => string = () => {
    throw new Error("GenOfficePptxFormatRuntime.setPathResolver not wired");
  };
  private readFile: (path: string) => Promise<Buffer> = async (path) =>
    (await import("node:fs/promises")).readFile(path);

  setPathResolver(resolver: (ref: string) => string): void {
    this.pathResolver = resolver;
    this.fallback.setPathResolver(resolver);
  }

  static deckOf(context: ArtifactContext): PptxDeck | undefined {
    return context.enrichment.get(DECK_KEY) as PptxDeck | undefined;
  }
}

/**
 * §101 honest resident estimates for engine read models. Element/block
 * records dominate; slide XML strings in the deck add the rest. These are
 * estimates (not exact heap walks) but within the right order of magnitude,
 * unlike the flat 8KB default.
 */
function estimateDeckBytes(deck: PptxDeck, sourceBytes: number): number {
  let elements = 0;
  for (const slide of deck.slides) elements += slide.elements.length;
  // ~1.2KB per element record (geometry + text body) + raw XML retained.
  return elements * 1_200 + sourceBytes * 0.35 + 65_536;
}

function estimateDocBytes(parsed: ParsedDocx, sourceBytes: number): number {
  // ~0.6KB per block (text + props); XML strings retained for patching.
  return parsed.blocks.length * 600 + sourceBytes * 0.5 + 65_536;
}

export class GenOfficeDocxFormatRuntime implements FormatRuntime {
  readonly format = "docx" as const;
  /** §127: degraded fallback when the engine cannot make sense of the bytes. */
  private readonly fallback = new BasicFormatRuntime("docx");

  async initialize(): Promise<void> {}

  async createArtifactContext(input: ArtifactBuildInput): Promise<ArtifactContext> {
    try {
      return await this.createFromEngine(input);
    } catch {
      await this.fallback.initialize();
      return this.fallback.createArtifactContext(input);
    }
  }

  private async createFromEngine(input: ArtifactBuildInput): Promise<ArtifactContext> {
    const path = this.pathResolver(input.artifactRef);
    const fingerprint = await fileFingerprint(path);
    if (input.signal?.aborted) throw new DOMException("aborted", "AbortError");
    // Zero-copy: wrap the read buffer's ArrayBuffer (no duplicate payload in memory).
    const buffer = await this.readFile(path);
    const bytes = new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
    const parsed = await parseDocx(bytes);
    if (parsed.blocks.length === 0) {
      throw new Error("engine produced no blocks");
    }
    if (input.signal?.aborted) throw new DOMException("aborted", "AbortError");

    const context: ArtifactContext = {
      artifactRef: input.artifactRef,
      version: { artifactRef: input.artifactRef, fingerprint },
      format: "docx",
      consistency: input.consistency,
      profile: "full",
      rendererVersion: GENOFFICE_RENDERER_VERSION,
      lastAccessAt: Date.now(),
      estimatedResidentBytes: estimateDocBytes(parsed, buffer.byteLength),
      enrichment: new Map<string, unknown>([
        ["engine", "genoffice"],
        ["blockCount", parsed.blocks.length],
        [DOC_KEY, parsed]
      ])
    };
    return context;
  }

  async trimMemory(level: MemoryTrimLevel): Promise<void> {
    await this.fallback.trimMemory(level);
  }

  async dispose(): Promise<void> {
    await this.fallback.dispose();
  }

  private pathResolver: (ref: string) => string = () => {
    throw new Error("GenOfficeDocxFormatRuntime.setPathResolver not wired");
  };
  private readFile: (path: string) => Promise<Buffer> = async (path) =>
    (await import("node:fs/promises")).readFile(path);

  setPathResolver(resolver: (ref: string) => string): void {
    this.pathResolver = resolver;
    this.fallback.setPathResolver(resolver);
  }

  static docOf(context: ArtifactContext): ParsedDocx | undefined {
    return context.enrichment.get(DOC_KEY) as ParsedDocx | undefined;
  }
}

/** Flatten docx blocks into preview-friendly rows (depth-first, text-bearing). */
export function flattenDocxBlocks(blocks: DocxBlock[], depth = 0, out: Array<{ index: number; style?: string; text: string }> = []): Array<{ index: number; style?: string; text: string }> {
  for (const block of blocks) {
    if (typeof block.text === "string" && block.text.trim().length > 0) {
      out.push({ index: out.length, style: block.style ?? block.type, text: block.text });
    }
    if (block.children && depth < 8) {
      flattenDocxBlocks(block.children, depth + 1, out);
    }
  }
  return out;
}
