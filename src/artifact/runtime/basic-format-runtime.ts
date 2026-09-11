/**
 * BasicFormatRuntime (§26): process-wide per-format runtime used until a
 * GenOffice-backed runtime is vendored (§146). Builds metadata/index-level
 * ArtifactContexts: ZIP index + [Content_Types] + format outline presence.
 * Shares one scanner cache across documents.
 */

import type {
  ArtifactBuildInput,
  ArtifactContext,
  FormatRuntime,
  MemoryTrimLevel
} from "../../contracts/artifact.js";
import type { OfficeFormat } from "../../contracts/ids.js";
import { fileFingerprint } from "../../support/fsx.js";
import { hasContentTypes, readZipIndex, type ZipIndex } from "../scanner/zip.js";

export const BASIC_RENDERER_VERSION = "basic-1";

export interface BasicArtifactState {
  zipIndex: ZipIndex;
  contentTypesPresent: boolean;
}

const stateByContext = new WeakMap<ArtifactContext, BasicArtifactState>();

export class BasicFormatRuntime implements FormatRuntime {
  private initialized = false;
  private readonly parsedIndexCache = new Map<string, ZipIndex>();
  private static readonly MAX_CACHED_INDEXES = 16;

  constructor(readonly format: OfficeFormat) {}

  async initialize(): Promise<void> {
    this.initialized = true;
  }

  async createArtifactContext(input: ArtifactBuildInput): Promise<ArtifactContext> {
    if (!this.initialized) await this.initialize();
    if (input.signal?.aborted) throw new DOMException("aborted", "AbortError");

    const path = this.pathResolver(input.artifactRef);
    const fingerprint = await fileFingerprint(path);
    const zipIndex = await readZipIndex(path);
    if (input.signal?.aborted) throw new DOMException("aborted", "AbortError");

    const state: BasicArtifactState = {
      zipIndex,
      contentTypesPresent: hasContentTypes(zipIndex)
    };
    this.cacheIndex(input.artifactRef, state);

    const context: ArtifactContext = {
      artifactRef: input.artifactRef,
      version: { artifactRef: input.artifactRef, fingerprint },
      format: this.format,
      consistency: input.consistency,
      rendererVersion: BASIC_RENDERER_VERSION,
      lastAccessAt: Date.now(),
      enrichment: new Map<string, unknown>([
        ["entryCount", zipIndex.entries.length],
        ["contentTypesPresent", state.contentTypesPresent]
      ])
    };
    stateByContext.set(context, state);
    return context;
  }

  /** Shared parse: later consumers of the same artifact reuse the ZIP index. */
  sharedIndex(artifactRef: string): ZipIndex | undefined {
    return this.parsedIndexCache.get(artifactRef);
  }

  private cacheIndex(artifactRef: string, state: BasicArtifactState): void {
    if (this.parsedIndexCache.size >= BasicFormatRuntime.MAX_CACHED_INDEXES) {
      const oldest = this.parsedIndexCache.keys().next().value;
      if (oldest !== undefined) this.parsedIndexCache.delete(oldest);
    }
    this.parsedIndexCache.set(artifactRef, state.zipIndex);
  }

  async trimMemory(level: MemoryTrimLevel): Promise<void> {
    if (level === "light" && this.parsedIndexCache.size > 4) {
      const excess = this.parsedIndexCache.size - 4;
      let i = 0;
      for (const key of this.parsedIndexCache.keys()) {
        if (i++ >= excess) break;
        this.parsedIndexCache.delete(key);
      }
      return;
    }
    if (level !== "light") this.parsedIndexCache.clear();
  }

  async dispose(): Promise<void> {
    this.parsedIndexCache.clear();
    this.initialized = false;
  }

  private pathResolver: (ref: string) => string = () => {
    throw new Error("BasicFormatRuntime.setPathResolver not wired");
  };

  setPathResolver(resolver: (ref: string) => string): void {
    this.pathResolver = resolver;
  }

  static stateOf(context: ArtifactContext): BasicArtifactState | undefined {
    return stateByContext.get(context);
  }
}
