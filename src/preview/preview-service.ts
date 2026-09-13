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
import { Scheduler, type ScheduledHandle, type SchedulerJob } from "../runtime/scheduler/scheduler.js";
import { SCHEDULER_PRIORITIES, type SchedulerPriorityName } from "../contracts/scheduler.js";
import type { ArtifactContext, ArtifactLease } from "../contracts/artifact.js";
import type { PreviewOutline } from "../contracts/preview.js";
import {
  GenOfficeDocxFormatRuntime,
  GenOfficePptxFormatRuntime,
  flattenDocxBlocks
} from "../artifact/runtime/genoffice-format-runtime.js";
import { elementText, buildRenderSlide } from "../vendor/genoffice/wrapper.js";
import { isRenderSlideLike, renderSlideToSvg } from "../vendor/genoffice/render-svg.js";
import { renderDocxOutline, renderPptxOutline, renderXlsxOutline } from "./outline-renderers.js";

/** Shared scheduler handle surface used by the in-flight render entry. */
interface SharedHandle {
  cancel(): void;
  promote(priority: SchedulerPriorityName): void;
}

/**
 * Shared in-flight render (round 10 reopen hardening):
 *  - `consumers` refcounts live preview consumers; shared queued work is
 *    cancelled only when the LAST consumer detaches (registry parity);
 *  - `highestIndex` remembers the best consumer priority — handles created
 *    AFTER a join (e.g. still inside registry.acquire) promote to it
 *    immediately, so inheritance no longer depends on handle creation order;
 *  - every consumer performs its own registry.acquire — joiners promote a
 *    still-queued build through the registry's own last-consumer semantics.
 */
interface InflightRender {
  promise: Promise<PreviewModel>;
  handles: SharedHandle[];
  highestIndex: number;
  consumers: Set<string>;
  /** Round 10 reopen: the seed registry acquire carries THIS controller's
   *  signal — aborted only when the last consumer leaves. A personal
   *  consumer abort must never poison the shared build for later joiners. */
  lifecycle: AbortController;
}

/**
 * Round 10: canonical scope key — cache/dedup identity includes the window.
 * Encoded as a fixed-order JSON tuple: values keep their delimiters, so
 * `{sheet:"A|max=1"}` and `{sheet:"A",maxEntries:1}` can never collide.
 */
export function normalizeScopeKey(scope?: PreviewScope): string {
  if (!scope) return "";
  const loc = scope.location ?? {};
  return JSON.stringify([
    loc.sheet ?? null,
    loc.slide ?? null,
    loc.block ?? null,
    loc.range ? [loc.range.fromRow, loc.range.fromCol, loc.range.toRow, loc.range.toCol] : null,
    scope.maxEntries ?? null
  ]);
}

/** Build an outline from a GenOffice-parsed context, or undefined when absent.
 *  Round 10: the engine model honors the requested scope window. */
function outlineFromGenOffice(context: ArtifactContext, scope?: PreviewScope): PreviewOutline | undefined {
  if (context.format === "pptx") {
    const deck = GenOfficePptxFormatRuntime.deckOf(context);
    if (!deck) return undefined;
    // Round 10 reopen: outline and SVG derive their slide window from the
    // SAME svgWindowOf clamp — an out-of-range anchor yields "last slide"
    // in BOTH views, never an empty outline next to a last-page SVG.
    const maxSlides = Math.min(200, Math.max(1, scope?.maxEntries ?? 200));
    const { from, count } = svgWindowOf(deck.slides.length, scope, maxSlides);
    return {
      kind: "pptx",
      slides: deck.slides.slice(from, from + count).map((slide, i) => ({
        index: from + i + 1,
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

/**
 * The scoped SVG slide window (pure math, exported for tests): outline and
 * SVG must always derive from the SAME window — `visual + slide 20` used to
 * render slides 1–6 while the outline showed 20+.
 */
export function svgWindowOf(
  slideCount: number,
  scope: PreviewScope | undefined,
  max: number
): { from: number; count: number } {
  const anchor = Math.max(0, Math.floor((scope?.location?.slide ?? 1) - 1));
  const from = Math.min(anchor, Math.max(0, slideCount - 1));
  const count = Math.max(0, Math.min(max, scope?.maxEntries ?? max, slideCount - from));
  return { from, count };
}

/** Headless visual render: slide draw lists → standalone SVG strings (§P7).
 *  Round 10: honors the scoped slide window (same slice as the outline). */
function svgSlidesFromGenOffice(context: ArtifactContext, scope?: PreviewScope): string[] | undefined {
  if (context.format !== "pptx") return undefined;
  const deck = GenOfficePptxFormatRuntime.deckOf(context);
  if (!deck) return undefined;
  const { from, count } = svgWindowOf(deck.slides.length, scope, 6);
  const out: string[] = [];
  for (const [i, slide] of deck.slides.slice(from, from + count).entries()) {
    try {
      const drawList = buildRenderSlide(slide, deck.size, { fitWidthPx: 960, slideNo: from + i + 1 });
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
    options?: {
      sheet?: string;
      /** 1-based inclusive window origin+extent (round 10). */
      range?: { fromRow: number; toRow: number; fromCol: number; toCol: number };
      maxRows?: number;
      maxCols?: number;
      maxSheets?: number;
    }
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
    const index = SCHEDULER_PRIORITIES.indexOf(PRIORITY_MAP[priority]);
    if (index < entry.highestIndex) {
      entry.highestIndex = index;
      for (const handle of entry.handles) handle.promote(PRIORITY_MAP[priority]);
    }
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
    const consumerId = `preview:${request.requestId}`;
    const existing = this.inflightDedup.get(cacheKeyString);
    const first = !existing;
    const shared: InflightRender = existing ?? {
      promise: undefined as never,
      handles: [],
      highestIndex: SCHEDULER_PRIORITIES.indexOf(PRIORITY_MAP[request.priority]),
      consumers: new Set<string>(),
      lifecycle: new AbortController()
    };
    if (first) {
      this.inflightDedup.set(cacheKeyString, shared);
      // The seed acquire belongs to the SHARED ENTRY, never to a personal
      // consumer signal (round 10 reopen): the first consumer aborting
      // mid-build must not poison the shared render later joiners rely on.
      // lifecycle aborts only when the LAST consumer leaves.
      const seed = this.registry.acquire({
        artifactRef: request.artifactRef,
        format,
        consistency: "optimistic",
        profile: request.visual ? "full" : "metadata",
        priority: PRIORITY_MAP[request.priority],
        consumer: `${consumerId}#seed`,
        signal: shared.lifecycle.signal
      });
      shared.promise = this.runSharedRender(shared, request, format, path, cacheKeyString, fingerprint, scopeKey, seed);
      // Guard: if every awaiter raced away on a personal abort, a later
      // shared rejection must not surface as unhandled.
      void shared.promise.catch(() => undefined);
    }

    // Every consumer (first included) registers itself. A detaching consumer
    // removes only ITSELF; the shared queued work and seed build are
    // cancelled when the LAST consumer leaves — matching registry semantics.
    shared.consumers.add(consumerId);
    this.promoteInflight(shared, request.priority);

    // A personal abort cancels only THIS consumer's wait (race rejection) —
    // never the shared render others still rely on. The listener is removed
    // on every settle path (success/error/abort), so a long-lived signal
    // never accumulates closures.
    let rejectPersonal: ((error: Error) => void) | undefined;
    const personal = new Promise<never>((_, reject) => {
      rejectPersonal = reject;
    });
    void personal.catch(() => undefined); // race loser stays quiet
    let onAbort: (() => void) | undefined;
    if (request.signal) {
      onAbort = () => {
        shared.consumers.delete(consumerId);
        if (shared.consumers.size === 0) {
          shared.lifecycle.abort();
          for (const handle of shared.handles) handle.cancel();
        }
        rejectPersonal?.(new DOMException("consumer aborted", "AbortError"));
      };
      if (!request.signal.aborted) {
        request.signal.addEventListener("abort", onAbort, { once: true });
      }
    }
    const awaited: Promise<PreviewModel> = request.signal
      ? Promise.race([shared.promise, personal])
      : shared.promise;

    // Joiners acquire the registry context themselves: the acquire JOINS the
    // same build (no duplicate parse) and promotes a still-queued build to
    // this consumer's priority — inheritance works even while the shared
    // render is still inside its own registry.acquire (no handles yet).
    const ownLease = first
      ? undefined
      : await this.registry.acquire({
          artifactRef: request.artifactRef,
          format,
          consistency: "optimistic",
          profile: request.visual ? "full" : "metadata",
          priority: PRIORITY_MAP[request.priority],
          consumer: consumerId,
          signal: request.signal
        });
    try {
      if (request.signal?.aborted && onAbort) onAbort(); // pre-aborted
      return await awaited;
    } finally {
      ownLease?.release();
      shared.consumers.delete(consumerId);
      if (onAbort && request.signal) request.signal.removeEventListener("abort", onAbort);
    }
  }

  private async runSharedRender(
    entry: InflightRender,
    request: PreviewRequest,
    format: OfficeFormat,
    path: string,
    cacheKeyString: string,
    fingerprint: { size: bigint; mtimeNs: bigint; fileId?: string },
    scopeKey: string,
    seedLease: Promise<ArtifactLease>
  ): Promise<PreviewModel> {
    /** Submit shared work under the CURRENT best priority; later joins keep
     *  promoting it via entry.handles. A submission racing the LAST consumer
     *  detaching (e.g. the fallback renderer after a cancelled sidecar read)
     *  cancels itself immediately — no orphaned queued work. */
    const submitShared = <T>(job: Omit<SchedulerJob<T>, "priority">): ScheduledHandle<T> => {
      const priority = SCHEDULER_PRIORITIES[entry.highestIndex]!;
      const handle = this.scheduler.submit<T>({ ...job, priority });
      entry.handles.push(handle);
      if (entry.consumers.size === 0) handle.cancel();
      return handle;
    };
    try {
      const lease = await seedLease;
      try {
        // §146 L6: prefer the GenOffice engine model already attached to the
        // context (deck / blocks); fall back to the streaming zip renderers.
        let outline = outlineFromGenOffice(lease.context, request.scope);
        // §30: XLSX goes through the Rust sidecar when present (bounded
        // viewport, workbook memory stays out of this process). Round 9/10:
        // the native call runs as a Scheduler job under the SAME priority
        // ladder; the dedicated single-flight permit matches the sidecar's
        // one native thread (general io/native budgets would let 4+ previews
        // enter the serial FIFO ahead of a later visible one). The full range
        // (origin + extent, 1-based) is forwarded — a resized top-left window
        // is not a range.
        if (!outline && format === "xlsx" && this.xlsxSidecar) {
          const sidecar = this.xlsxSidecar;
          const range = request.scope?.location?.range;
          const handle = submitShared({
            label: `preview-xlsx-sidecar:${request.artifactRef}`,
            resources: { io: 1, xlsxSidecar: 1 },
            run: async (signal) => {
              if (signal.aborted) throw new DOMException("cancelled", "AbortError");
              return await sidecar.previewWindow(path, {
                sheet: request.scope?.location?.sheet,
                range,
                maxRows: request.scope?.maxEntries ?? (range ? range.toRow - range.fromRow + 1 : undefined),
                maxCols: range ? range.toCol - range.fromCol + 1 : undefined
              });
            }
          });
          outline = await handle.promise
            .then((sheets): PreviewOutline => ({
              kind: "xlsx",
              sheets: sheets.map((sheet) => ({
                name: sheet.name,
                rowCount: sheet.rowCount ?? sheet.window.length,
                // Exact ONLY when native metadata provided a real extent;
                // window.length alone is a lower bound.
                rowCountExact: sheet.rowCount !== undefined,
                window: sheet.window
              }))
            }))
            .catch(() => undefined);
        }
        if (!outline) {
          // Everyone left while the sidecar attempt was being cancelled —
          // do not fall back into fresh work for abandoned consumers.
          if (entry.consumers.size === 0) throw new DOMException("cancelled", "AbortError");
          const handle = submitShared({
            label: `preview-render:${request.artifactRef}`,
            resources: { io: 1 },
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
          outline = await handle.promise;
        }
        // Round 10 (P1): headless SVG generation is real CPU/render work —
        // it no longer runs unadmitted inside the shared promise.
        let svgSlides: string[] | undefined;
        if (request.visual) {
          const context = lease.context;
          const handle = submitShared({
            label: `preview-svg:${request.artifactRef}`,
            resources: { render: 1, cpu: 1 },
            run: async (signal) => {
              if (signal.aborted) throw new DOMException("cancelled", "AbortError");
              return svgSlidesFromGenOffice(context, request.scope) ?? [];
            }
          });
          const svgs = await handle.promise.catch(() => undefined);
          svgSlides = svgs && svgs.length > 0 ? svgs : undefined;
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
          svgSlides
        };
        if (this.cache.admit(JSON.stringify(model).length)) {
          this.cache.set(cacheKeyString, model);
        }
        return model;
      } finally {
        lease.release();
      }
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
