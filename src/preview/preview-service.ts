/**
 * PreviewService (§11, §18, §157): high-frequency read path. Creates no
 * DocumentSession, no WriterLease, no candidate, never spawns OfficeCLI
 * (PERF-01/02/03), never strong-hashes before first paint (PERF-07).
 * Optimistic consistency: fingerprint before → render → fingerprint after;
 * mismatched renders are recomputed once.
 */

import type {
  PreviewModel,
  PreviewRequest,
  PreviewResult
} from "../contracts/preview.js";
import type { PreviewCacheKey } from "../contracts/artifact.js";
import type { OfficeFormat } from "../contracts/ids.js";
import { fileFingerprint, fingerprintKey } from "../support/fsx.js";
import { ArtifactRegistry } from "../artifact/registry/artifact-registry.js";
import { ByteBudgetCache } from "../artifact/cache/byte-budget-cache.js";
import type { ArtifactStore } from "../artifact/store/artifact-store.js";
import type { Scheduler } from "../runtime/scheduler/scheduler.js";
import type { ArtifactContext } from "../contracts/artifact.js";
import type { PreviewOutline } from "../contracts/preview.js";
import {
  GenOfficeDocxFormatRuntime,
  GenOfficePptxFormatRuntime,
  flattenDocxBlocks
} from "../artifact/runtime/genoffice-format-runtime.js";
import { elementText, buildRenderSlide } from "../vendor/genoffice/wrapper.js";
import { isRenderSlideLike, renderSlideToSvg } from "../vendor/genoffice/render-svg.js";
import { renderDocxOutline, renderPptxOutline, renderXlsxOutline } from "./outline-renderers.js";

/** Build an outline from a GenOffice-parsed context, or undefined when absent. */
function outlineFromGenOffice(context: ArtifactContext): PreviewOutline | undefined {
  if (context.format === "pptx") {
    const deck = GenOfficePptxFormatRuntime.deckOf(context);
    if (!deck) return undefined;
    return {
      kind: "pptx",
      slides: deck.slides.slice(0, 200).map((slide, i) => ({
        index: i + 1,
        shapes: slide.elements
          .slice(0, 50)
          .map((element) => ({ name: element.name, text: elementText(element) || undefined }))
      }))
    };
  }
  if (context.format === "docx") {
    const doc = GenOfficeDocxFormatRuntime.docOf(context);
    if (!doc) return undefined;
    return { kind: "docx", blocks: flattenDocxBlocks(doc.blocks).slice(0, 400) };
  }
  return undefined;
}

/** Headless visual render: slide draw lists → standalone SVG strings (§P7). */
function svgSlidesFromGenOffice(context: ArtifactContext, maxSlides: number): string[] | undefined {
  if (context.format !== "pptx") return undefined;
  const deck = GenOfficePptxFormatRuntime.deckOf(context);
  if (!deck) return undefined;
  const out: string[] = [];
  for (const [i, slide] of deck.slides.slice(0, maxSlides).entries()) {
    try {
      const drawList = buildRenderSlide(slide, deck.size, { fitWidthPx: 960, slideNo: i + 1 });
      if (isRenderSlideLike(drawList)) out.push(renderSlideToSvg(drawList));
    } catch {
      // Visual rendering is best-effort per slide; the outline still serves.
    }
  }
  return out.length > 0 ? out : undefined;
}

const PREVIEW_PROFILE = "outline-v1";
const FONT_ENVIRONMENT_ID = "system-default";

const PRIORITY_MAP = {
  visible: "VISIBLE_PREVIEW",
  prefetch: "PREFETCH",
  background: "BACKGROUND_INDEX"
} as const;

export interface XlsxSidecarPort {
  previewWindow(
    path: string,
    options?: { maxRows?: number; maxCols?: number; maxSheets?: number }
  ): Promise<Array<{ name: string; window: string[][]; rowCount?: number }>>;
}

export class PreviewService {
  /** §91: visual preview cache — highly recyclable, byte budgeted. */
  readonly cache = new ByteBudgetCache<PreviewModel>(
    "VisualPreviewCache",
    128 * 1024 * 1024,
    (model) => JSON.stringify(model).length
  );
  private inflightDedup = new Map<string, Promise<PreviewModel>>();

  constructor(
    private readonly store: ArtifactStore,
    private readonly registry: ArtifactRegistry,
    private readonly scheduler: Scheduler,
    private readonly xlsxSidecar?: XlsxSidecarPort
  ) {}

  async preview(request: PreviewRequest): Promise<PreviewResult> {
    const started = Date.now();
    const format = this.store.formatOf(request.artifactRef);
    const path = this.store.resolvePath(request.artifactRef);

    const before = await fileFingerprint(path);
    const cacheKeyString = previewCacheKeyString(request.artifactRef, fingerprintKey(before), request.visual);
    const cached = this.cache.get(cacheKeyString);
    if (cached) {
      return { requestId: request.requestId, model: cached, retried: false, elapsedMs: Date.now() - started };
    }

    let model = await this.render(request, format, path, cacheKeyString, before);
    let retried = false;

    // Optimistic consistency (§18): re-probe after render; retry once on change.
    const after = await fileFingerprint(path);
    if (fingerprintKey(after) !== fingerprintKey(before)) {
      retried = true;
      const retryKey = previewCacheKeyString(request.artifactRef, fingerprintKey(after), request.visual);
      model = await this.render(request, format, path, retryKey, after);
    }

    return { requestId: request.requestId, model, retried, elapsedMs: Date.now() - started };
  }

  private async render(
    request: PreviewRequest,
    format: OfficeFormat,
    path: string,
    cacheKeyString: string,
    fingerprint: { size: bigint; mtimeNs: bigint; fileId?: string }
  ): Promise<PreviewModel> {
    const inflight = this.inflightDedup.get(cacheKeyString);
    if (inflight) return inflight;

    const job = (async () => {
      // Registry acquire: dedupes parses across preview/open/edit consumers (§23).
      const lease = await this.registry.acquire({
        artifactRef: request.artifactRef,
        format,
        consistency: "optimistic",
        // P0-7: quick previews stay light (ZIP/index); visual previews
        // upgrade to the full engine read model.
        profile: request.visual ? "full" : "metadata",
        priority: PRIORITY_MAP[request.priority],
        consumer: `preview:${request.requestId}`
      });
      try {
        // §146 L6: prefer the GenOffice engine model already attached to the
        // context (deck / blocks); fall back to the streaming zip renderers.
        let outline = outlineFromGenOffice(lease.context);
        // §30: XLSX goes through the Rust sidecar when present (bounded
        // viewport, workbook memory stays out of this process).
        if (!outline && format === "xlsx" && this.xlsxSidecar) {
          outline = await this.xlsxSidecar
            .previewWindow(path)
            .then((sheets): PreviewOutline => ({
              kind: "xlsx",
              sheets: sheets.map((sheet) => ({
                name: sheet.name,
                rowCount: sheet.rowCount ?? sheet.window.length,
                window: sheet.window
              }))
            }))
            .catch(() => undefined);
        }
        if (!outline) {
          outline = await this.scheduler
            .submit({
              label: `preview-render:${request.artifactRef}`,
              priority: PRIORITY_MAP[request.priority],
              run: async () => {
                switch (format) {
                  case "docx":
                    return await renderDocxOutline(path);
                  case "xlsx":
                    return await renderXlsxOutline(path);
                  case "pptx":
                    return await renderPptxOutline(path);
                }
              }
            })
            .promise;
        }
        const cacheKey: PreviewCacheKey = {
          contentHash: fingerprintKey(fingerprint),
          rendererVersion: lease.context.rendererVersion,
          fontEnvironmentId: FONT_ENVIRONMENT_ID,
          previewProfile: PREVIEW_PROFILE
        };
        const model: PreviewModel = {
          format,
          outline,
          cacheKey,
          fingerprintAtRender: fingerprintKey(fingerprint),
          // SVG rendering requires the full engine model (visual profile only).
          svgSlides: request.visual ? svgSlidesFromGenOffice(lease.context, 6) : undefined
        };
        if (this.cache.admit(JSON.stringify(model).length)) {
          this.cache.set(cacheKeyString, model);
        }
        return model;
      } finally {
        lease.release();
      }
    })();

    this.inflightDedup.set(cacheKeyString, job);
    try {
      return await job;
    } finally {
      this.inflightDedup.delete(cacheKeyString);
    }
  }

  /** Invalidate preview cache entries for an artifact (revision change). */
  invalidate(artifactRef: string): void {
    for (const [key] of Array.from(this.cache.entries())) {
      if (key.startsWith(`${artifactRef}|`)) this.cache.delete(key);
    }
  }

  stats(): { cacheSize: number; cacheBytes: number; hits: number; misses: number } {
    return {
      cacheSize: this.cache.length,
      cacheBytes: this.cache.currentBytes,
      hits: this.cache.stats.hits,
      misses: this.cache.stats.misses
    };
  }

  trimAll(): number {
    return this.cache.trimAll();
  }
}

function previewCacheKeyString(ref: string, fpKey: string, visual?: boolean): string {
  return `${ref}|${fpKey}${visual ? "|v" : ""}`;
}
