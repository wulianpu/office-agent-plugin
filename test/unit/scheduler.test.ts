import { describe, expect, it, vi } from "vitest";
import { Scheduler } from "../../src/runtime/scheduler/scheduler.js";

const delay = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

describe("Scheduler (§110–§112)", () => {
  it("runs INTERACTIVE before BACKGROUND_INDEX regardless of submit order", async () => {
    const scheduler = new Scheduler({ maxConcurrent: 1 });
    const order: string[] = [];
    // Occupy the single lane so both candidates queue first.
    const blocker = scheduler.submit({
      label: "blocker",
      priority: "INTERACTIVE",
      run: () => delay(20)
    });
    const bg = scheduler.submit({
      label: "bg",
      priority: "BACKGROUND_INDEX",
      run: async () => {
        order.push("bg");
      }
    });
    const fg = scheduler.submit({
      label: "fg",
      priority: "INTERACTIVE",
      run: async () => {
        order.push("fg");
      }
    });
    await Promise.all([blocker.promise, bg.promise, fg.promise]);
    expect(order).toEqual(["fg", "bg"]);
  });

  it("drops stale jobs instead of delivering results (backpressure §112, §43)", async () => {
    const scheduler = new Scheduler({ maxConcurrent: 1 });
    const onStale = vi.fn();
    const handle = scheduler.submit<{ value: number }>({
      label: "stale-job",
      priority: "VERIFICATION",
      identity: { sessionId: "s1", sessionEpoch: 1, artifactRef: "a" },
      current: () => ({ sessionEpoch: 2 }), // epoch moved → stale
      run: async () => {
        await delay(20);
        return { value: 42 };
      },
      onStale
    });
    const result = await handle.promise;
    expect(result).toBeUndefined();
    expect(onStale).toHaveBeenCalled();
    expect(scheduler.stats.staleDropped).toBe(1);
  });

  it("cancelWhere removes queued preview work for a superseded revision", async () => {
    const scheduler = new Scheduler({ maxConcurrent: 1 });
    // Occupy the single lane.
    const blocker = scheduler.submit({
      label: "blocker",
      priority: "INTERACTIVE",
      run: () => delay(40)
    });
    const queued = scheduler.submit({
      label: "preview-render:art1",
      priority: "PREFETCH",
      run: () => delay(1)
    });
    const removed = scheduler.cancelWhere((job) => job.label.startsWith("preview-render:art1"));
    expect(removed).toBe(1);
    await expect(queued.promise).rejects.toThrow();
    await blocker.promise;
  });
});
