/**
 * Round 9 (issue #3, P1): the XLSX sidecar preview runs through the Scheduler
 * priority ladder — a queued BACKGROUND read must not be served before a
 * later VISIBLE preview when both wait on the same single slot.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { PreviewService } from "../../src/preview/preview-service.js";
import { ArtifactRegistry } from "../../src/artifact/registry/artifact-registry.js";
import { Scheduler } from "../../src/runtime/scheduler/scheduler.js";
import { ResourceGovernor } from "../../src/runtime/resources/resource-governor.js";
import type { ArtifactStore } from "../../src/artifact/store/artifact-store.js";
import type { FormatRuntime } from "../../src/contracts/artifact.js";

let dir: string;
let bgPath: string;
let visPath: string;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "preview-prio-"));
  bgPath = join(dir, "background.xlsx");
  visPath = join(dir, "visible.xlsx");
  await writeFile(bgPath, Buffer.alloc(64, 3));
  await writeFile(visPath, Buffer.alloc(64, 5));
});

afterAll(async () => {
  await rm(dir, { recursive: true, force: true }).catch(() => undefined);
});

/** Instant metadata-context runtime: previews stay on the sidecar path. */
function xlsxRuntime(): FormatRuntime {
  return {
    format: "xlsx",
    async createArtifactContext(input) {
      return {
        artifactRef: input.artifactRef,
        version: { artifactRef: input.artifactRef, fingerprint: { size: 1n, mtimeNs: 1n } },
        format: "xlsx",
        consistency: input.consistency,
        rendererVersion: "test",
        lastAccessAt: Date.now(),
        enrichment: new Map()
      };
    },
    initialize: async () => undefined,
    trimMemory: async () => undefined,
    dispose: async () => undefined
  };
}

async function waitFor(condition: () => boolean, deadlineMs = 5_000): Promise<void> {
  const deadline = Date.now() + deadlineMs;
  while (!condition()) {
    if (Date.now() > deadline) throw new Error("waitFor: condition not met in time");
    await new Promise((resolve) => setTimeout(resolve, 1));
  }
}

describe("PreviewService sidecar scheduling (round 9, issue #3)", () => {
  it("visible xlsx previews overtake queued background sidecar reads", async () => {
    const scheduler = new Scheduler({ maxConcurrent: 1 });
    const registry = new ArtifactRegistry();
    registry.registerRuntime(xlsxRuntime());
    registry.setPathResolver((ref) => (ref === "art-bg" ? bgPath : visPath));

    const sidecarOrder: string[] = [];
    const service = new PreviewService(
      {
        formatOf: () => "xlsx",
        resolvePath: (ref: string) => (ref === "art-bg" ? bgPath : visPath)
      } as unknown as ArtifactStore,
      registry,
      scheduler,
      {
        previewWindow: async (path: string) => {
          sidecarOrder.push(path);
          return [{ name: "Sheet1", window: [["a"]], rowCount: 1 }];
        }
      }
    );

    // Occupy the single scheduler slot, then queue background BEFORE visible.
    let release!: () => void;
    const blocker = scheduler.submit({
      label: "blocker",
      priority: "INTERACTIVE",
      resources: { io: 1 },
      run: () => new Promise<void>((resolve) => (release = resolve))
    });
    void blocker.promise.catch(() => undefined);
    await waitFor(() => scheduler.runningCount === 1);

    const background = service.preview({ requestId: "req-bg", artifactRef: "art-bg", priority: "background" });
    const visible = service.preview({ requestId: "req-vis", artifactRef: "art-vis", priority: "visible" });
    await waitFor(() => scheduler.queueDepth === 2);

    release(); // free the slot — the ladder must pick visible first
    await waitFor(() => sidecarOrder.length === 2);

    // The visible preview entered the native sidecar BEFORE the earlier-
    // queued background read.
    expect(sidecarOrder).toEqual([visPath, bgPath]);
    const [bgModel, visModel] = await Promise.all([background, visible]);
    expect(bgModel.model.outline.kind).toBe("xlsx");
    expect(visModel.model.outline.kind).toBe("xlsx");
  });

  it("production defaults (8 lanes / 4 io): the single-flight sidecar permit keeps visible out of the native FIFO tail (round 10)", async () => {
    // Issue #4 acceptance: with generic io permits, 4 background previews
    // could ALL enter the one-threaded sidecar FIFO before a later visible
    // request. The dedicated xlsx-sidecar permit (=1) serializes admission.
    const governor = new ResourceGovernor(); // DEFAULT_BUDGETS: io 4, sidecar 1
    const scheduler = new Scheduler({ governor }); // maxConcurrent 8
    const registry = new ArtifactRegistry();
    registry.registerRuntime(xlsxRuntime());
    const files: Array<{ ref: string; path: string }> = [];
    for (let i = 0; i < 5; i++) {
      const path = join(dir, `sf-${i}.xlsx`);
      await writeFile(path, Buffer.alloc(32, i + 1));
      files.push({ ref: `sf-${i}`, path });
    }
    registry.setPathResolver((ref) => files.find((f) => f.ref === ref)!.path);

    let inFlight = 0;
    let maxConcurrent = 0;
    const entered: string[] = [];
    const waiters: Array<() => void> = [];
    const service = new PreviewService(
      {
        formatOf: () => "xlsx",
        resolvePath: (ref: string) => files.find((f) => f.ref === ref)!.path
      } as unknown as ArtifactStore,
      registry,
      scheduler,
      {
        previewWindow: async (path: string) => {
          inFlight++;
          maxConcurrent = Math.max(maxConcurrent, inFlight);
          entered.push(path);
          await new Promise<void>((resolve) => waiters.push(resolve));
          inFlight--;
          return [{ name: "Sheet1", window: [["x"]], rowCount: 1 }];
        }
      }
    );

    const backgroundPreviews = files.slice(0, 4).map((f) =>
      service.preview({ requestId: `bg-${f.ref}`, artifactRef: f.ref, priority: "background" })
    );
    // Exactly one background enters the native call; the rest stay queued on
    // the sidecar permit.
    await waitFor(() => entered.length === 1);

    const visiblePromise = service.preview({
      requestId: "vis",
      artifactRef: files[4]!.ref,
      priority: "visible"
    });
    await waitFor(() => scheduler.queueDepth >= 4); // 3 bg + 1 vis queued

    // Drain the native FIFO one entry at a time.
    const drain = () => waiters.splice(0)[0]?.();
    drain();
    await waitFor(() => entered.length === 2);
    // The VISIBLE preview entered SECOND — not after all four background
    // reads (it would be 5th under the old io-only admission).
    expect(entered[1]).toBe(files[4]!.path);
    for (let i = 0; i < 10 && waiters.length >= 0; i++) {
      drain();
      if (entered.length >= 5) break;
      await new Promise((resolve) => setTimeout(resolve, 1));
    }
    await Promise.all([...backgroundPreviews, visiblePromise]);
    expect(maxConcurrent).toBe(1); // single-flight into the native process
  });

  it("dedup priority inheritance: a visible joiner promotes the shared background render (round 10)", async () => {
    // Order: background(A) queues → prefetch(B) queues → visible joins A.
    // Promotion must lift A's render above B's PREFETCH — without it B
    // (PREFETCH outranks BACKGROUND) would run first.
    const scheduler = new Scheduler({ maxConcurrent: 1 });
    const registry = new ArtifactRegistry();
    registry.registerRuntime(xlsxRuntime());
    const aPath = join(dir, "inherit-a.xlsx");
    const bPath = join(dir, "inherit-b.xlsx");
    await writeFile(aPath, Buffer.alloc(24, 1));
    await writeFile(bPath, Buffer.alloc(24, 2));
    registry.setPathResolver((ref) => (ref === "art-a" ? aPath : bPath));

    const entries: string[] = [];
    const service = new PreviewService(
      {
        formatOf: () => "xlsx",
        resolvePath: (ref: string) => (ref === "art-a" ? aPath : bPath)
      } as unknown as ArtifactStore,
      registry,
      scheduler,
      {
        previewWindow: async (path: string) => {
          entries.push(path);
          return [{ name: "Sheet1", window: [["y"]], rowCount: 1 }];
        }
      }
    );

    let release!: () => void;
    const blocker = scheduler.submit({
      label: "blocker",
      priority: "INTERACTIVE",
      resources: { io: 1 },
      run: () => new Promise<void>((resolve) => (release = resolve))
    });
    void blocker.promise.catch(() => undefined);
    await waitFor(() => scheduler.runningCount === 1);

    const backgroundA = service.preview({ requestId: "bg-a", artifactRef: "art-a", priority: "background" });
    await waitFor(() => scheduler.queueDepth === 1); // A's sidecar job queued
    const prefetchB = service.preview({ requestId: "pf-b", artifactRef: "art-b", priority: "prefetch" });
    await waitFor(() => scheduler.queueDepth === 2);
    // Visible joins the SAME artifact/scope → dedup hit, promotes the shared job.
    const visibleA = service.preview({ requestId: "vis-a", artifactRef: "art-a", priority: "visible" });
    // The join is one stat-await behind; let it land BEFORE releasing the
    // lane so the promotion is observably in the queue order.
    await new Promise((resolve) => setTimeout(resolve, 25));
    expect(scheduler.queueDepth).toBe(2); // both still queued — nothing raced ahead

    release();
    await Promise.all([backgroundA, prefetchB, visibleA]);
    // A was promoted to VISIBLE and dispatches before B's PREFETCH.
    expect(entries).toEqual([aPath, bPath]);
    // Still exactly ONE native read for artifact A (dedup preserved).
    expect(entries.filter((p) => p === aPath)).toHaveLength(1);
  });

  it("consumer abort dequeues a queued sidecar preview before it reaches the native process (round 10)", async () => {
    const scheduler = new Scheduler({ maxConcurrent: 1 });
    const registry = new ArtifactRegistry();
    registry.registerRuntime(xlsxRuntime());
    const path = join(dir, "abort-queued.xlsx");
    await writeFile(path, Buffer.alloc(24, 9));
    registry.setPathResolver(() => path);

    const entered: string[] = [];
    let release!: () => void;
    const service = new PreviewService(
      { formatOf: () => "xlsx", resolvePath: () => path } as unknown as ArtifactStore,
      registry,
      scheduler,
      {
        previewWindow: async (p: string) => {
          entered.push(p);
          await new Promise<void>((resolve) => (release = resolve));
          return [{ name: "Sheet1", window: [["z"]], rowCount: 1 }];
        }
      }
    );

    const blocker = scheduler.submit({
      label: "blocker",
      priority: "INTERACTIVE",
      run: () => new Promise<void>(() => undefined) // stays until cancelled
    });
    void blocker.promise.catch(() => undefined);
    await waitFor(() => scheduler.runningCount === 1);

    const controller = new AbortController();
    const queued = service.preview({
      requestId: "abort-me",
      artifactRef: "art-abort",
      priority: "background",
      signal: controller.signal
    });
    void queued.catch(() => undefined);
    await waitFor(() => scheduler.queueDepth === 1);

    controller.abort(); // superseded while queued
    await new Promise((resolve) => setImmediate(resolve));
    await new Promise((resolve) => setImmediate(resolve));
    expect(scheduler.queueDepth).toBe(0); // dequeued before native execution
    blocker.cancel();
    await expect(queued).rejects.toMatchObject({ name: "AbortError" });
    expect(entered).toHaveLength(0); // never reached the native process
  });
});
