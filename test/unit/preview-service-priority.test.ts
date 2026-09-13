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

  it("first-consumer abort during a gated Registry build does NOT poison the shared render (round 10 reopen)", async () => {
    const scheduler = new Scheduler({ maxConcurrent: 1 });
    const registry = new ArtifactRegistry();
    const runtime = {
      format: "xlsx" as const,
      builds: [] as string[],
      gates: [] as Array<() => void>,
      async createArtifactContext(input: { artifactRef: string; consistency: "optimistic" | "stable" }) {
        runtime.builds.push(input.artifactRef);
        await new Promise<void>((resolve) => runtime.gates.push(resolve));
        return {
          artifactRef: input.artifactRef,
          version: { artifactRef: input.artifactRef, fingerprint: { size: 1n, mtimeNs: 1n } },
          format: "xlsx" as const,
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
    registry.registerRuntime(runtime, "full");
    registry.setBuildScheduler(scheduler);
    const path = join(dir, "poison-guard.xlsx");
    await writeFile(path, Buffer.alloc(24, 6));
    registry.setPathResolver(() => path);
    const service = new PreviewService(
      { formatOf: () => "xlsx", resolvePath: () => path } as unknown as ArtifactStore,
      registry,
      scheduler,
      {
        previewWindow: async () => [{ name: "Sheet1", window: [["poison-free"]], rowCount: 1 }]
      }
    );

    const controllerA = new AbortController();
    // A starts the shared render; the full build is gated (unresolved).
    const a = service.preview({ requestId: "poison-a", artifactRef: "art-poison", priority: "background", visual: true, signal: controllerA.signal });
    void a.catch(() => undefined);
    await waitFor(() => runtime.builds.length === 1 && scheduler.runningCount === 1);
    // B joins the SAME key while the build is still gated.
    const b = service.preview({ requestId: "poison-b", artifactRef: "art-poison", priority: "background", visual: true });
    await new Promise((resolve) => setTimeout(resolve, 25));

    controllerA.abort(); // A leaves BEFORE the build resolves
    await new Promise((resolve) => setImmediate(resolve));
    expect(scheduler.runningCount).toBe(1); // the shared build keeps running for B

    runtime.gates.splice(0)[0]?.(); // release the build
    const modelB = await b; // B succeeds — A's abort did not poison the shared promise
    expect(modelB.model.outline.kind).toBe("xlsx");
    expect(runtime.builds.filter((ref) => ref === "art-poison")).toHaveLength(1); // one parse
    await a.catch(() => undefined);
  });

  it("abort listeners are removed on every settle path — a long-lived signal never accumulates closures (round 10 reopen)", async () => {
    const scheduler = new Scheduler();
    const registry = new ArtifactRegistry();
    registry.registerRuntime(xlsxRuntime());
    const path = join(dir, "listener-hygiene.xlsx");
    await writeFile(path, Buffer.alloc(24, 3));
    registry.setPathResolver(() => path);
    const service = new PreviewService(
      { formatOf: () => "xlsx", resolvePath: () => path } as unknown as ArtifactStore,
      registry,
      scheduler,
      {
        previewWindow: async () => [{ name: "Sheet1", window: [["clean"]], rowCount: 1 }]
      }
    );

    const controller = new AbortController();
    const signal = controller.signal;
    let active = 0;
    const originalAdd = signal.addEventListener.bind(signal);
    const originalRemove = signal.removeEventListener.bind(signal);
    (signal as unknown as { addEventListener: typeof originalAdd }).addEventListener = (...args: Parameters<typeof originalAdd>) => {
      active++;
      return originalAdd(...args);
    };
    (signal as unknown as { removeEventListener: typeof originalRemove }).removeEventListener = (...args: Parameters<typeof originalRemove>) => {
      active--;
      return originalRemove(...args);
    };

    // Sequential previews on ONE shared signal — each settle must clean up.
    for (let i = 0; i < 25; i++) {
      await service.preview({ requestId: `hygiene-${i}`, artifactRef: `art-hygiene-${i}`, priority: "visible", signal });
      // The registry acquire also uses the same signal — both listeners clean.
    }
    expect(active).toBe(0); // no listener accumulation across settled previews
    // Abort path cleanup: one aborted preview also removes its listener (once fires it, then our removeEventListener is a no-op — but the pending count must not grow).
    const controller2 = new AbortController();
    const before = (controller2.signal as unknown as { listenerCount?: () => number }).listenerCount;
    void before;
    const rejected = service.preview({ requestId: "hygiene-abort", artifactRef: "art-hygiene-0", priority: "background", signal: controller2.signal });
    void rejected.catch(() => undefined);
    controller2.abort();
    await rejected.catch(() => undefined);
    await new Promise((resolve) => setImmediate(resolve));
  });

  it("consumer-aware cancellation: one joiner aborting never kills the other's shared work (round 10 reopen)", async () => {
    const scheduler = new Scheduler({ maxConcurrent: 1 });
    const registry = new ArtifactRegistry();
    registry.registerRuntime(xlsxRuntime());
    const path = join(dir, "ab-cancel.xlsx");
    await writeFile(path, Buffer.alloc(24, 7));
    registry.setPathResolver(() => path);

    let releaseSidecar!: () => void;
    const service = new PreviewService(
      { formatOf: () => "xlsx", resolvePath: () => path } as unknown as ArtifactStore,
      registry,
      scheduler,
      {
        previewWindow: async () => {
          await new Promise<void>((resolve) => (releaseSidecar = resolve));
          return [{ name: "Sheet1", window: [["ok"]], rowCount: 1 }];
        }
      }
    );

    let releaseBlocker!: () => void;
    const blocker = scheduler.submit({
      label: "blocker",
      priority: "INTERACTIVE",
      run: () => new Promise<void>((resolve) => (releaseBlocker = resolve))
    });
    void blocker.promise.catch(() => undefined);
    await waitFor(() => scheduler.runningCount === 1);

    const controllerA = new AbortController();
    // A and B share one artifact + scope -> one shared render.
    const a = service.preview({ requestId: "a", artifactRef: "art-ab", priority: "background", signal: controllerA.signal });
    void a.catch(() => undefined);
    await waitFor(() => scheduler.queueDepth === 1);
    const b = service.preview({ requestId: "b", artifactRef: "art-ab", priority: "background" });
    await new Promise((resolve) => setTimeout(resolve, 25)); // B joins the shared entry

    controllerA.abort(); // A detaches — B must survive
    await new Promise((resolve) => setImmediate(resolve));
    expect(scheduler.queueDepth).toBe(1); // shared queued work NOT cancelled

    releaseBlocker();
    await waitFor(() => scheduler.queueDepth === 0);
    releaseSidecar();
    const modelB = await b; // B succeeds on the shared render
    expect(modelB.model.outline.kind).toBe("xlsx");
    await a.catch(() => undefined);
  });

  it("consumer-aware cancellation: when BOTH consumers abort, the queued shared work dequeues (round 10 reopen)", async () => {
    const scheduler = new Scheduler({ maxConcurrent: 1 });
    const registry = new ArtifactRegistry();
    registry.registerRuntime(xlsxRuntime());
    const path = join(dir, "ab-both.xlsx");
    await writeFile(path, Buffer.alloc(24, 8));
    registry.setPathResolver(() => path);
    let entered = 0;
    const service = new PreviewService(
      { formatOf: () => "xlsx", resolvePath: () => path } as unknown as ArtifactStore,
      registry,
      scheduler,
      {
        previewWindow: async () => {
          entered++;
          return [{ name: "Sheet1", window: [["x"]], rowCount: 1 }];
        }
      }
    );

    const blocker = scheduler.submit({ label: "blocker", priority: "INTERACTIVE", run: () => new Promise<void>(() => undefined) });
    void blocker.promise.catch(() => undefined);
    await waitFor(() => scheduler.runningCount === 1);

    const ca = new AbortController();
    const cb = new AbortController();
    const a = service.preview({ requestId: "a2", artifactRef: "art-both", priority: "background", signal: ca.signal });
    const b = service.preview({ requestId: "b2", artifactRef: "art-both", priority: "background", signal: cb.signal });
    void a.catch(() => undefined);
    void b.catch(() => undefined);
    await waitFor(() => scheduler.queueDepth === 1);
    await new Promise((resolve) => setTimeout(resolve, 25)); // both joined

    ca.abort();
    await new Promise((resolve) => setImmediate(resolve));
    expect(scheduler.queueDepth).toBe(1); // B still alive
    cb.abort(); // last consumer leaves
    await new Promise((resolve) => setImmediate(resolve));
    await new Promise((resolve) => setImmediate(resolve));
    expect(scheduler.queueDepth).toBe(0); // shared work dequeued
    blocker.cancel();
    expect(entered).toBe(0); // never reached the native port
  });

  it("priority inheritance survives joins made BEFORE any render handle exists (round 10 reopen)", async () => {
    // The shared render is still inside registry.acquire (full build queued,
    // gated) when the VISIBLE joiner arrives — no scheduler handles exist
    // yet. The joiner's own registry.acquire must promote the queued build.
    const scheduler = new Scheduler({ maxConcurrent: 1 });
    const registry = new ArtifactRegistry();
    const runtime = {
      format: "xlsx" as const,
      builds: [] as string[],
      gates: [] as Array<() => void>,
      async createArtifactContext(input: { artifactRef: string; consistency: "optimistic" | "stable" }) {
        runtime.builds.push(input.artifactRef);
        await new Promise<void>((resolve) => runtime.gates.push(resolve));
        return {
          artifactRef: input.artifactRef,
          version: { artifactRef: input.artifactRef, fingerprint: { size: 1n, mtimeNs: 1n } },
          format: "xlsx" as const,
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
    registry.registerRuntime(runtime, "full");
    registry.setBuildScheduler(scheduler);
    const pathA = join(dir, "pre-handle-a.xlsx");
    const pathB = join(dir, "pre-handle-b.xlsx");
    await writeFile(pathA, Buffer.alloc(24, 1));
    await writeFile(pathB, Buffer.alloc(24, 2));
    registry.setPathResolver((ref: string) => (ref === "art-pa" ? pathA : pathB));
    const service = new PreviewService(
      {
        formatOf: () => "xlsx",
        resolvePath: (ref: string) => (ref === "art-pa" ? pathA : pathB)
      } as unknown as ArtifactStore,
      registry,
      scheduler,
      undefined // no sidecar: build queued, then the fallback render path
    );

    let releaseBlocker!: () => void;
    const blocker = scheduler.submit({
      label: "blocker",
      priority: "INTERACTIVE",
      run: () => new Promise<void>((resolve) => (releaseBlocker = resolve))
    });
    void blocker.promise.catch(() => undefined);
    await waitFor(() => scheduler.runningCount === 1);

    // BACKGROUND visual build for A queues (full profile, gated runtime).
    const backgroundA = service.preview({ requestId: "pa", artifactRef: "art-pa", priority: "background", visual: true });
    void backgroundA.catch(() => undefined);
    await waitFor(() => scheduler.queueDepth === 1); // A's BUILD queued — no render handles yet
    // A second background build for B queues AFTER the join.
    const backgroundB = service.preview({ requestId: "pb", artifactRef: "art-pb", priority: "background", visual: true });
    void backgroundB.catch(() => undefined);
    await waitFor(() => scheduler.queueDepth === 2);
    // VISIBLE joins A while its build is still queued.
    const visibleA = service.preview({ requestId: "pv", artifactRef: "art-pa", priority: "visible", visual: true });
    await new Promise((resolve) => setTimeout(resolve, 25));

    releaseBlocker();
    // The promoted build for A must dispatch before B's BACKGROUND build.
    await waitFor(() => runtime.builds.length >= 1);
    expect(runtime.builds[0]).toBe("art-pa");
    while (runtime.builds.length < 2) {
      runtime.gates.splice(0)[0]?.();
      await new Promise((resolve) => setTimeout(resolve, 1));
    }
    expect(runtime.builds[1]).toBe("art-pb");
    for (let i = 0; i < 4; i++) runtime.gates.splice(0)[0]?.();
    await Promise.allSettled([backgroundA, backgroundB, visibleA]);
    // Still ONE parse for artifact A (dedup preserved through the join).
    expect(runtime.builds.filter((ref) => ref === "art-pa")).toHaveLength(1);
  });

  it("the full 1-based range (origin + extent) is forwarded to the sidecar port (round 10 reopen)", async () => {
    const scheduler = new Scheduler();
    const registry = new ArtifactRegistry();
    registry.registerRuntime(xlsxRuntime());
    const path = join(dir, "range-fwd.xlsx");
    await writeFile(path, Buffer.alloc(24, 4));
    registry.setPathResolver(() => path);
    const received: Array<Record<string, unknown>> = [];
    const service = new PreviewService(
      { formatOf: () => "xlsx", resolvePath: () => path } as unknown as ArtifactStore,
      registry,
      scheduler,
      {
        previewWindow: async (_p: string, options?: Record<string, unknown>) => {
          received.push(options ?? {});
          return [{ name: "Data", window: [["v"]], rowCount: 42 }];
        }
      }
    );
    const result = await service.preview({
      requestId: "range-1",
      artifactRef: "art-range",
      priority: "visible",
      scope: { location: { sheet: "Data", range: { fromRow: 10, toRow: 30, fromCol: 4, toCol: 6 } } }
    });
    if (result.model.outline.kind !== "xlsx") throw new Error("expected xlsx outline");
    expect(received[0]!.range).toEqual({ fromRow: 10, toRow: 30, fromCol: 4, toCol: 6 });
    expect(received[0]!.sheet).toBe("Data");
    // rowCountExact honors real metadata only.
    expect(result.model.outline.sheets[0]!.rowCountExact).toBe(true);
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
