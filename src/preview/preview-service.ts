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
  PreviewResult,
  PreviewScope
} from "../contracts/preview.js";
import type { PreviewCacheKey } from "../contracts/artifact.js";
import type { OfficeFormat } from "../contracts/ids.js";
import { fileFingerprint, fingerprintKey } from "../support/fsx.js";
import { ArtifactRegistry } from "../artifact/registry/artifact-registry.js";
import { ByteBudgetCache } from "../artifact/cache/byte-budget-cache.js";
import type { ArtifactStore } from "../artifact/store/artifact-store.js";
import type { Scheduler } from "../runtime/scheduler/scheduler.js";
import type { SchedulerPriorityName } from "../contracts/scheduler.js";
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

/** Shared in-flight render: joiners promote the scheduler handles upward. */
interface InflightRender {
  promise: Promise<PreviewModel>;
  handles: Array<{ promote(priority: SchedulerPriorityName): void }>;
}

/**
 * Round 10: canonical scope key — cache/dedup identity includes the window.
 * Field order is fixed so equal scopes always produce equal keys.
 */
export function normalizeScopeKey(scope?: PreviewScope): string {
  if (!scope) return "";
  const loc = scope.location ?? {};
  const parts: string[] = [];
  if (loc.sheet !== undefined) parts.push(`sheet=${loc.sheet}`);
  if (loc.slide !== undefined) parts.push(`slide=${loc.slide}`);
  if (loc.block !== undefined) parts.push(`block=${loc.block}`);
  if (loc.range) {
    parts.push(`range=${loc.range.fromRow},${loc.range.fromCol}-${loc.range.toRow},${loc.range.toCol}`);
  }
  if (scope.maxEntries !== undefined) parts.push(`max=${scope.maxEntries}`);
  return parts.join("|");
}

/** Consumer abort dequeues queued scheduler work before it reaches the
 *  sidecar/engine; running work finishes (its result is simply unused). */
function wireConsumerAbort(signal: AbortSignal | undefined, handle: { cancel(): void }): void {
  if (!signal) return;
  if (signal.aborted) {
    handle.cancel();
    return;
  }
  signal.addEventListener("abort", () => handle.cancel(), { once: true });
}

/** Build an outline from a GenOffice-parsed context, or undefined when absent.
 *  Round 10: the engine model honors the requested scope window. */
function outlineFromGenOffice(context: ArtifactContext, scope?: PreviewScope): PreviewOutline | undefined {
  if (context.format === "pptx") {
    const deck = GenOfficePptxFormatRuntime.deckOf(context);
    if (!deck) return undefined;
    const anchor = Math.max(0, Math.floor((scope?.location?.slide ?? 1) - 1));
    const maxSlides = Math.min(200, Math.max(1, scope?.maxEntries ?? 200));
    return {
      kind: "pptx",
      slides: deck.slides.slice(anchor, anchor + maxSlides).map((slide, i) => ({
        index: anchor + i + 1,
        shapes: slide.elements
          .slice(0, 50)
          .map((element) => ({ name: element.name, text: elementText(element) || undefined }))
      }))
    };
  }
  if (context.format === "docx") {
    const doc = GenOfficeDocxFormatRuntime.docOf(context);
    if (!doc) return undefined;
    const anchor = Math.max(0, Math.floor(scope?.location?.block ?? 0));
    const maxBlocks = Math.min(400, Math.max(1, scope?.maxEntries ?? 400));
    const flat = flattenDocxBlocks(doc.blocks);
    return { kind: "docx", blocks: flat.slice(anchor, anchor + maxBlocks).map((b, i) => ({ ...b, index: anchor + i })) };
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
    options?: { sheet?: string; maxRows?: number; maxCols?: number; maxSheets?: number }
  ): Promise<Array<{ name: string; window: string[][]; rowCount?: number }>>;
}

export class PreviewService {
  /** §91: visual preview cache — highly recyclable, byte budgeted. */
  readonly cache = new ByteBudgetCache<PreviewModel>(
    "VisualPreviewCache",
    128 * 1024 * 1024,
    (model) => JSON.stringify(model).length
  );
  private inflightDedup = new Map<string, InflightRender>();

  /**
   * Round 10 priority inheritance: a later VISIBLE consumer joining an
   * in-flight BACKGROUND render promotes the shared scheduler handles
   * (never downward) instead of duplicating the parse work.
   */
  private promoteInflight(entry: InflightRender, priority: keyof typeof PRIORITY_MAP): void {
    for (const handle of entry.handles) handle.promote(PRIORITY_MAP[priority]);
  }

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
    // Round 10: the normalized scope is part of every cache/dedup identity —
    // two different windows of one artifact never share a result.
    const scopeKey = normalizeScopeKey(request.scope);
    const cacheKeyString = previewCacheKeyString(request.artifactRef, fingerprintKey(before), request.visual, scopeKey);
    const cached = this.cache.get(cacheKeyString);
    if (cached) {
      return { requestId: request.requestId, model: cached, retried: false, elapsedMs: Date.now() - started };
    }

    let model = await this.render(request, format, path, cacheKeyString, before, scopeKey);
    let retried = false;

    // Optimistic consistency (§18): re-probe after render; retry once on change.
    const after = await fileFingerprint(path);
    if (fingerprintKey(after) !== fingerprintKey(before)) {
      retried = true;
      const retryKey = previewCacheKeyString(request.artifactRef, fingerprintKey(after), request.visual, scopeKey);
      model = await this.render(request, format, path, retryKey, after, scopeKey);
    }

    return { requestId: request.requestId, model, retried, elapsedMs: Date.now() - started };
  }

  private async render(
    request: PreviewRequest,
    format: OfficeFormat,
    path: string,
    cacheKeyString: string,
    fingerprint: { size: bigint; mtimeNs: bigint; fileId?: string },
    scopeKey: string
  ): Promise<PreviewModel> {
    const inflight = this.inflightDedup.get(cacheKeyString);
    if (inflight) {
      // Round 10 (P1): join the shared work — no duplicate parse — but
      // inherit the consumer's priority (promote handles upward only).
      this.promoteInflight(inflight, request.priority);
      return inflight.promise;
    }

    const handles: Array<{ promote(priority: SchedulerPriorityName): void }> = [];
    const entry: InflightRender = { promise: undefined as never, handles };
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
        let outline = outlineFromGenOffice(lease.context, request.scope);
        // §30: XLSX goes through the Rust sidecar when present (bounded
        // viewport, workbook memory stays out of this process). Round 9/10:
        // the native call runs as a Scheduler job under the SAME priority
        // ladder; the dedicated single-flight permit matches the sidecar's
        // one native thread (general io/native budgets would let 4+ previews
        // enter the serial FIFO ahead of a later visible one).
        if (!outline && format === "xlsx" && this.xlsxSidecar) {
          const sidecar = this.xlsxSidecar;
          const range = request.scope?.location?.range;
          const handle = this.scheduler.submit({
            label: `preview-xlsx-sidecar:${request.artifactRef}`,
            priority: PRIORITY_MAP[request.priority],
            resources: { io: 1, xlsxSidecar: 1 },
            run: async (signal) => {
              if (signal.aborted) throw new DOMException("cancelled", "AbortError");
              return await sidecar.previewWindow(path, {
                sheet: request.scope?.location?.sheet,
                maxRows: request.scope?.maxEntries ?? (range ? range.toRow - range.fromRow + 1 : undefined),
                maxCols: range ? range.toCol - range.fromCol + 1 : undefined
              });
            }
          });
          handles.push(handle);
          wireConsumerAbort(request.signal, handle);
          outline = await handle.promise
            .then((sheets): PreviewOutline => ({
              kind: "xlsx",
              sheets: sheets.map((sheet) => ({
                name: sheet.name,
                rowCount: sheet.rowCount ?? sheet.window.length,
                rowCountExact: true, // sidecar sheet metadata is authoritative
                window: sheet.window
              }))
            }))
            .catch(() => undefined);
        }
        if (!outline) {
          const handle = this.scheduler.submit({
            label: `preview-render:${request.artifactRef}`,
            priority: PRIORITY_MAP[request.priority],
            run: async () => {
              switch (format) {
                case "docx":
                  return await renderDocxOutline(path, request.scope);
                case "xlsx":
                  return await renderXlsxOutline(path, request.scope);
                case "pptx":
                  return await renderPptxOutline(path, request.scope);
              }
            }
          });
          handles.push(handle);
          wireConsumerAbort(request.signal, handle);
          outline = await handle.promise;
        }
        const cacheKey: PreviewCacheKey = {
          contentHash: fingerprintKey(fingerprint),
          rendererVersion: lease.context.rendererVersion,
          fontEnvironmentId: FONT_ENVIRONMENT_ID,
          previewProfile: PREVIEW_PROFILE,
          scope: scopeKey || undefined
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

    entry.promise = job;
    this.inflightDedup.set(cacheKeyString, entry);
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

function previewCacheKeyString(ref: string, fpKey: string, visual?: boolean, scopeKey?: string): string {
  return `${ref}|${fpKey}${visual ? "|v" : ""}${scopeKey ? `|s:${scopeKey}` : ""}`;
}
