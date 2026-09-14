/**
 * Issue #4 (reopen, P1-high/P1): memoryBytes is a real global admission
 * boundary; cache resident bytes join the same ledger; full builds reserve
 * their transient peak via estimatedBytes.
 */

import { describe, expect, it } from "vitest";
import { ResourceGovernor, DEFAULT_BUDGETS } from "../../src/runtime/resources/resource-governor.js";
import { ByteBudgetCache } from "../../src/artifact/cache/byte-budget-cache.js";

describe("governor memory admission (issue #4)", () => {
  it("unprotected memory reservations are REFUSED beyond the global budget; protected bypasses admission but counts", () => {
    const governor = new ResourceGovernor({
      ...DEFAULT_BUDGETS,
      memoryBytes: 1000
    });
    const ok = governor.tryAcquire({ classes: { memory: 1 }, bytes: 600, label: "a" });
    expect(ok).toBeDefined();
    // 600 + 600 > 1000 → refused (memory-only, no other class involved).
    expect(governor.tryAcquire({ classes: { memory: 1 }, bytes: 600, label: "b" })).toBeUndefined();
    // Protected writer state bypasses admission but still counts.
    const prot = governor.tryAcquire({ classes: { memory: 1 }, bytes: 300, label: "p", protected: true });
    expect(prot).toBeDefined();
    expect(governor.memoryUsage().used).toBe(900);
    ok?.release();
    prot?.release();
    expect(governor.memoryUsage().used).toBe(0);
    // Freed → admission works again.
    expect(governor.tryAcquire({ classes: { memory: 1 }, bytes: 900, label: "c" })).toBeDefined();
  });

  it("acquire()'s retry loop trims registered caches when memory alone blocks admission", async () => {
    const governor = new ResourceGovernor({ ...DEFAULT_BUDGETS, memoryBytes: 1000 });
    const cache = new ByteBudgetCache<{ blob: string }>("ArtifactCache", 800, (v) => v.blob.length, { maxItemShareOfBudget: 1 });
    // Simulate pre-existing resident cache bytes in the same ledger.
    cache.set("big", { blob: "x".repeat(700) });
    governor.reportCacheBytes(700);
    let trimmed = false;
    governor.registerTrimTarget({
      rank: 1,
      label: "test-cache",
      trim: async () => {
        trimmed = true;
        const removed = cache.currentBytes;
        cache.clear(); // clear reports -removed via its reporter if wired
        governor.reportCacheBytes(-removed);
      }
    });
    // 700 held + 500 wanted > 1000: acquire must retry → trim frees → admit.
    const lease = await governor.acquire({ classes: { memory: 1 }, bytes: 500, label: "needs-trim" });
    expect(trimmed).toBe(true);
    expect(lease).toBeDefined();
    expect(governor.memoryUsage().used).toBe(500);
    lease.release();
  });

  it("cache growth crossing the pressure boundary triggers the eviction ladder automatically", async () => {
    const governor = new ResourceGovernor({ ...DEFAULT_BUDGETS, memoryBytes: 1000 });
    let trims = 0;
    governor.registerTrimTarget({
      rank: 1,
      label: "auto-trim",
      trim: async () => {
        trims++;
      }
    });
    expect(governor.getPressure().level).toBe("normal");
    governor.reportCacheBytes(1800); // over budget → critical
    await new Promise((resolve) => setTimeout(resolve, 50));
    expect(governor.getPressure().level).toBe("critical");
    expect(trims).toBeGreaterThan(0); // ladder fired without any acquire()
  });
});

describe("ByteBudgetCache delta reporting (issue #4)", () => {
  it("set/overwrite/evict/delete/clear report exact NET deltas", () => {
    const deltas: number[] = [];
    const cache = new ByteBudgetCache<{ blob: string }>("ArtifactCache", 1000, (v) => v.blob.length, { maxItemShareOfBudget: 1 }, (d) =>
      deltas.push(d)
    );
    cache.set("a", { blob: "x".repeat(100) });
    cache.set("b", { blob: "y".repeat(200) });
    cache.set("a", { blob: "z".repeat(50) }); // overwrite: net -50
    cache.delete("b"); // -200
    cache.clear(); // -50
    expect(deltas).toEqual([100, 200, -50, -200, -50]);
    // Ledger nets to zero — no double charge, no underflow.
    expect(deltas.reduce((a, b) => a + b, 0)).toBe(0);
    expect(cache.currentBytes).toBe(0);
  });

  it("LRU eviction reports the net effect of one set()", () => {
    const deltas: number[] = [];
    const cache = new ByteBudgetCache<{ blob: string }>("ArtifactCache", 250, (v) => v.blob.length, { maxItemShareOfBudget: 1 }, (d) =>
      deltas.push(d)
    );
    cache.set("a", { blob: "x".repeat(100) });
    cache.set("b", { blob: "y".repeat(100) });
    cache.set("c", { blob: "z".repeat(100) }); // evicts a → net 0
    expect(deltas).toEqual([100, 100, 0]);
    expect(cache.currentBytes).toBe(200);
  });
});

describe("full-build transient peak reservation (issue #4 P1)", () => {
  it("a full build declares estimatedBytes and stays QUEUED when the global memory budget cannot admit it", async () => {
    const { ArtifactRegistry } = await import("../../src/artifact/registry/artifact-registry.js");
    const { Scheduler } = await import("../../src/runtime/scheduler/scheduler.js");
    const { mkdtemp, rm, writeFile } = await import("node:fs/promises");
    const { tmpdir } = await import("node:os");
    const { join } = await import("node:path");
    const dir = await mkdtemp(join(tmpdir(), "peak-res-"));
    try {
      const path = join(dir, "big.pptx");
      await writeFile(path, Buffer.alloc(64 * 1024)); // 64 KiB source → peak ≥ 16 MiB floor
      const governor = new ResourceGovernor({ ...DEFAULT_BUDGETS, memoryBytes: 8 * 1024 * 1024 });
      const scheduler = new Scheduler({ governor });
      const registry = new ArtifactRegistry();
      const runtime = {
        format: "pptx" as const,
        builds: 0,
        async createArtifactContext(input: { artifactRef: string; consistency: "optimistic" | "stable" }) {
          runtime.builds++;
          return {
            artifactRef: input.artifactRef,
            version: { artifactRef: input.artifactRef, fingerprint: { size: 1n, mtimeNs: 1n } },
            format: "pptx" as const,
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
      registry.setPathResolver(() => path);

      const acquired = registry.acquire({
        artifactRef: "ref-peak",
        format: "pptx",
        consistency: "optimistic",
        profile: "full",
        priority: "VISIBLE_PREVIEW",
        consumer: "peak-test"
      });
      const settled = await Promise.race([
        acquired.then(() => "admitted"),
        new Promise((resolve) => setTimeout(() => resolve("queued"), 300))
      ]);
      expect(settled).toBe("queued"); // 16 MiB floor > 8 MiB budget → stays queued
      expect(runtime.builds).toBe(0); // engine never started
      void acquired.catch(() => undefined);
    } finally {
      const { rm } = await import("node:fs/promises");
      await rm(dir, { recursive: true, force: true }).catch(() => undefined);
    }
  });
});
