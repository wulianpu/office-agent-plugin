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

/** Parsed models are attached to contexts via WeakMap (§25 progressive enrichment). */
const deckByContext = new WeakMap<ArtifactContext, PptxDeck>();
const docByContext = new WeakMap<ArtifactContext, ParsedDocx>();

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
    const bytes = new Uint8Array(await this.readFile(path));
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
      rendererVersion: GENOFFICE_RENDERER_VERSION,
      lastAccessAt: Date.now(),
      enrichment: new Map<string, unknown>([
        ["engine", "genoffice"],
        ["slideCount", opened.deck.slides.length],
        ["deckSize", opened.deck.size]
      ])
    };
    deckByContext.set(context, opened.deck);
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
    return deckByContext.get(context);
  }
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
    const bytes = new Uint8Array(await this.readFile(path));
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
      rendererVersion: GENOFFICE_RENDERER_VERSION,
      lastAccessAt: Date.now(),
      enrichment: new Map<string, unknown>([
        ["engine", "genoffice"],
        ["blockCount", parsed.blocks.length]
      ])
    };
    docByContext.set(context, parsed);
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
    return docByContext.get(context);
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
