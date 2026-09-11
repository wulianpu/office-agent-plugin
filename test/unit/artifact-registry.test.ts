import { describe, expect, it, vi, beforeAll, afterAll } from "vitest";
import { writeFile, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { ArtifactRegistry } from "../../src/artifact/registry/artifact-registry.js";
import type {
  ArtifactBuildInput,
  ArtifactContext,
  FormatRuntime
} from "../../src/contracts/artifact.js";

let fixtureDir: string;
let fixturePath: string;

beforeAll(async () => {
  fixtureDir = await mkdtemp(join(tmpdir(), "registry-test-"));
  fixturePath = join(fixtureDir, "artifact.pptx");
  await writeFile(fixturePath, Buffer.alloc(64, 1));
});

afterAll(async () => {
  await rm(fixtureDir, { recursive: true, force: true });
});

class ControllableRuntime implements FormatRuntime {
  readonly format = "pptx" as const;
  builds = 0;
  aborted = 0;
  /** When true, builds stay pending until the test resolves them via gate(). */
  pending = false;
  private gates: Array<() => void> = [];
  initialize = vi.fn(async () => undefined);
  trimMemory = vi.fn(async () => undefined);
  dispose = vi.fn(async () => undefined);

  async createArtifactContext(input: ArtifactBuildInput): Promise<ArtifactContext> {
    this.builds++;
    const context: ArtifactContext = {
      artifactRef: input.artifactRef,
      version: { artifactRef: input.artifactRef, fingerprint: { size: 1n, mtimeNs: BigInt(this.builds) } },
      format: "pptx",
      consistency: input.consistency,
      rendererVersion: "test",
      lastAccessAt: Date.now(),
      enrichment: new Map()
    };
    if (input.signal) {
      input.signal.addEventListener("abort", () => this.aborted++, { once: true });
    }
    if (!this.pending) return context;
    return new Promise<ArtifactContext>((resolve, reject) => {
      const onAbort = () => reject(new DOMException("aborted", "AbortError"));
      input.signal?.addEventListener("abort", onAbort, { once: true });
      this.gates.push(() => {
        input.signal?.removeEventListener("abort", onAbort);
        resolve(context);
      });
    });
  }

  gate(): void {
    for (const g of this.gates.splice(0)) g();
  }
}

/** Deterministic barrier: wait until a condition holds (stat timing varies). */
async function waitFor(condition: () => boolean): Promise<void> {
  for (let i = 0; i < 100 && !condition(); i++) {
    await new Promise((resolve) => setImmediate(resolve));
  }
}

const acquire = (registry: ArtifactRegistry, consumer: string, signal?: AbortSignal) =>
  registry.acquire({
    artifactRef: "a",
    format: "pptx",
    consistency: "optimistic",
    priority: "VISIBLE_PREVIEW",
    consumer,
    signal
  });

describe("ArtifactRegistry (§21–§24, PERF-04/05/06)", () => {
  it("deduplicates concurrent acquires into one build (1 build → 3 consumers)", async () => {
    const runtime = new ControllableRuntime();
    const registry = new ArtifactRegistry();
    registry.registerRuntime(runtime);
    registry.setPathResolver(() => fixturePath);

    const leases = await Promise.all([acquire(registry, "preview:1"), acquire(registry, "open:1"), acquire(registry, "edit:1")]);
    expect(runtime.builds).toBe(1);
    for (const lease of leases) lease.release();
  });

  it("a consumer aborting does not kill the shared build while others wait (PERF-06)", async () => {
    const runtime = new ControllableRuntime();
    runtime.pending = true;
    const registry = new ArtifactRegistry();
    registry.registerRuntime(runtime);
    registry.setPathResolver(() => fixturePath);

    const controllerA = new AbortController();
    const promiseA = acquire(registry, "preview:A", controllerA.signal);
    const promiseB = acquire(registry, "open:B");
    // Deterministic: both consumers must have joined the shared build.
    await waitFor(() => registry.leaseCount() >= 2);

    controllerA.abort(); // Preview cancelled; Open still needs the build.
    await expect(promiseA).rejects.toMatchObject({ name: "AbortError" });

    runtime.gate();
    const leaseB = await promiseB;
    expect(runtime.builds).toBe(1);
    expect(runtime.aborted).toBe(0); // shared build was never aborted
    leaseB.release();
  });

  it("aborts the build when the last consumer detaches, and later acquires rebuild fresh", async () => {
    const runtime = new ControllableRuntime();
    runtime.pending = true;
    const registry = new ArtifactRegistry();
    registry.registerRuntime(runtime);
    registry.setPathResolver(() => fixturePath);

    const controller = new AbortController();
    const promise = acquire(registry, "preview:solo", controller.signal);
    await waitFor(() => registry.leaseCount() >= 1); // ensure registration
    controller.abort();
    await expect(promise).rejects.toMatchObject({ name: "AbortError" });
    expect(runtime.aborted).toBe(1);
    expect(registry.inFlightCount()).toBe(0);

    // A later consumer gets a fresh build (not a poisoned one).
    runtime.pending = false;
    const lease = await acquire(registry, "preview:next");
    expect(runtime.builds).toBe(2);
    lease.release();
  });

  it("reuses a completed context without rebuilding (PERF-05)", async () => {
    const runtime = new ControllableRuntime();
    const registry = new ArtifactRegistry();
    registry.registerRuntime(runtime);
    registry.setPathResolver(() => fixturePath);

    const first = await acquire(registry, "c1");
    first.release();
    const second = await acquire(registry, "c2");
    second.release();
    expect(runtime.builds).toBe(1);
  });
});
