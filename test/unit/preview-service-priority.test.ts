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
    const runtime: FormatRuntime = {
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
    registry.registerRuntime(runtime);
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
});
