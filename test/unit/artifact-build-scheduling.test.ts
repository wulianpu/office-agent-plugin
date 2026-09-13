/**
 * Round 10 (issue #4, P1-high): ArtifactRegistry context builds enter the
 * shared Scheduler — a visible full-profile parse is not stuck behind queued
 * background builds; joiners promote the shared build; the last consumer's
 * abort dequeues a still-queued build.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { ArtifactRegistry } from "../../src/artifact/registry/artifact-registry.js";
import { Scheduler } from "../../src/runtime/scheduler/scheduler.js";
import type { ArtifactBuildInput, ArtifactContext, FormatRuntime } from "../../src/contracts/artifact.js";

let dir: string;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "build-sched-"));
});

afterAll(async () => {
  await rm(dir, { recursive: true, force: true }).catch(() => undefined);
});

async function waitFor(condition: () => boolean, deadlineMs = 5_000): Promise<void> {
  const deadline = Date.now() + deadlineMs;
  while (!condition()) {
    if (Date.now() > deadline) throw new Error("waitFor: condition not met in time");
    await new Promise((resolve) => setTimeout(resolve, 1));
  }
}

/** Gated full-profile runtime: builds stay pending until the test releases. */
class GatedRuntime implements FormatRuntime {
  readonly format = "pptx" as const;
  builds: string[] = [];
  private gates: Array<() => void> = [];

  async createArtifactContext(input: ArtifactBuildInput): Promise<ArtifactContext> {
    const tag = `${input.artifactRef}`;
    this.builds.push(tag);
    await new Promise<void>((resolve) => this.gates.push(resolve));
    return {
      artifactRef: input.artifactRef,
      version: { artifactRef: input.artifactRef, fingerprint: { size: 1n, mtimeNs: 1n } },
      format: "pptx",
      consistency: input.consistency,
      rendererVersion: "test",
      lastAccessAt: Date.now(),
      enrichment: new Map()
    };
  }

  releaseOne(): void {
    this.gates.splice(0)[0]?.();
  }

  initialize = async () => undefined;
  trimMemory = async () => undefined;
  dispose = async () => undefined;
}

async function makeFiles(count: number): Promise<Array<{ ref: string; path: string }>> {
  const files: Array<{ ref: string; path: string }> = [];
  for (let i = 0; i < count; i++) {
    const path = join(dir, `f${i}.pptx`);
    await writeFile(path, Buffer.alloc(32, i + 1));
    files.push({ ref: `ref-${i}`, path });
  }
  return files;
}

describe("ArtifactRegistry build scheduling (round 10, issue #4)", () => {
  it("visible build dispatches before queued background builds; joiners promote; abort dequeues", async () => {
    const files = await makeFiles(5);
    const scheduler = new Scheduler({ maxConcurrent: 1 });
    const registry = new ArtifactRegistry();
    const runtime = new GatedRuntime();
    registry.registerRuntime(runtime, "full");
    registry.setBuildScheduler(scheduler);
    registry.setPathResolver((ref) => files.find((f) => f.ref === ref)!.path);

    // Occupy the single lane so all builds queue first.
    let release!: () => void;
    const blocker = scheduler.submit({
      label: "blocker",
      priority: "INTERACTIVE",
      resources: { io: 1 },
      run: () => new Promise<void>((resolve) => (release = resolve))
    });
    void blocker.promise.catch(() => undefined);
    await waitFor(() => scheduler.runningCount === 1);

    // ref-0..2 background builds (distinct keys → distinct queued builds).
    const bg0 = registry.acquire({ artifactRef: files[0]!.ref, format: "pptx", consistency: "optimistic", profile: "full", priority: "BACKGROUND_INDEX", consumer: "c0" });
    const bg1 = registry.acquire({ artifactRef: files[1]!.ref, format: "pptx", consistency: "optimistic", profile: "full", priority: "BACKGROUND_INDEX", consumer: "c1" });
    const bg2 = registry.acquire({ artifactRef: files[2]!.ref, format: "pptx", consistency: "optimistic", profile: "full", priority: "BACKGROUND_INDEX", consumer: "c2" });
    // ref-3 joins as VISIBLE — promotion must lift it above every background.
    const bg3Promise = registry.acquire({ artifactRef: files[3]!.ref, format: "pptx", consistency: "optimistic", profile: "full", priority: "BACKGROUND_INDEX", consumer: "c3" });
    const join = registry.acquire({ artifactRef: files[3]!.ref, format: "pptx", consistency: "optimistic", profile: "full", priority: "VISIBLE_PREVIEW", consumer: "c3b" });
    // ref-4: queued background whose consumer aborts before dispatch.
    const abort = new AbortController();
    const aborted = registry.acquire({ artifactRef: files[4]!.ref, format: "pptx", consistency: "optimistic", profile: "full", priority: "BACKGROUND_INDEX", consumer: "c4", signal: abort.signal });
    const abortedGuard = aborted.catch(() => undefined);

    await waitFor(() => scheduler.queueDepth === 5);
    abort.abort(); // last consumer detaches while still queued
    await waitFor(() => scheduler.queueDepth === 4); // dequeued outright

    release();
    // Drain: VISIBLE ref-3 first, then background ref-0/1/2 in order.
    await waitFor(() => runtime.builds.length >= 1);
    expect(runtime.builds[0]).toBe(files[3]!.ref);
    while (runtime.builds.length < 4) {
      runtime.releaseOne();
      await new Promise((resolve) => setTimeout(resolve, 1));
    }
    expect(runtime.builds.slice(1)).toEqual([files[0]!.ref, files[1]!.ref, files[2]!.ref]);
    for (let i = 0; i < 4; i++) runtime.releaseOne();

    const leases = await Promise.all([bg0, bg1, bg2, bg3Promise, join]);
    for (const lease of leases) lease.release();
    await abortedGuard; // aborted acquire rejected (never built)
    expect(runtime.builds).not.toContain(files[4]!.ref);
    expect(registry.inFlightCount()).toBe(0);
  });

  it("same-key consumers still produce ONE engine parse (dedup preserved under the scheduler)", async () => {
    const files = await makeFiles(1);
    const scheduler = new Scheduler();
    const registry = new ArtifactRegistry();
    const runtime = new GatedRuntime();
    registry.registerRuntime(runtime, "full");
    registry.setBuildScheduler(scheduler);
    registry.setPathResolver((ref) => files.find((f) => f.ref === ref)!.path);

    const acquired = Promise.all([
      registry.acquire({ artifactRef: files[0]!.ref, format: "pptx", consistency: "optimistic", profile: "full", priority: "BACKGROUND_INDEX", consumer: "a" }),
      registry.acquire({ artifactRef: files[0]!.ref, format: "pptx", consistency: "optimistic", profile: "full", priority: "VISIBLE_PREVIEW", consumer: "b" }),
      registry.acquire({ artifactRef: files[0]!.ref, format: "pptx", consistency: "optimistic", profile: "full", priority: "VISIBLE_PREVIEW", consumer: "c" })
    ]);
    await waitFor(() => runtime.builds.length === 1);
    runtime.releaseOne(); // ungate the single shared build
    const leases = await acquired;
    expect(runtime.builds).toHaveLength(1); // one parse, three consumers
    for (const lease of leases) lease.release();
  });
});
