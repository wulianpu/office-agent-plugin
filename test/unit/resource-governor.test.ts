import { describe, expect, it, vi } from "vitest";
import {
  DEFAULT_BUDGETS,
  ResourceGovernor
} from "../../src/runtime/resources/resource-governor.js";

describe("ResourceGovernor (§105–§107)", () => {
  it("enforces native-process concurrency", async () => {
    const governor = new ResourceGovernor({
      ...DEFAULT_BUDGETS,
      nativeProcessConcurrent: 1
    });
    const first = governor.tryAcquire({ classes: { "native-process": 1 }, label: "p1" });
    expect(first).toBeDefined();
    const second = governor.tryAcquire({ classes: { "native-process": 1 }, label: "p2" });
    expect(second).toBeUndefined();
    first!.release();
    const third = governor.tryAcquire({ classes: { "native-process": 1 }, label: "p3" });
    expect(third).toBeDefined();
  });

  it("pressure escalates with utilization", () => {
    const governor = new ResourceGovernor({ ...DEFAULT_BUDGETS, ioConcurrent: 2 });
    expect(governor.getPressure().level).toBe("normal");
    const l1 = governor.tryAcquire({ classes: { io: 1 }, label: "io1" })!;
    expect(governor.getPressure().level).toBe("normal"); // 50% < 60%
    const l2 = governor.tryAcquire({ classes: { io: 1 }, label: "io2" })!;
    expect(governor.getPressure().level).toBe("critical"); // 100%
    l1.release();
    l2.release();
    expect(governor.getPressure().level).toBe("normal");
  });

  it("runs trim targets in ladder order under pressure", async () => {
    const order: string[] = [];
    const governor = new ResourceGovernor({ ...DEFAULT_BUDGETS, ioConcurrent: 1 });
    governor.registerTrimTarget({
      rank: 6,
      label: "cold-artifact-cache",
      trim: async () => {
        order.push("cold");
      }
    });
    governor.registerTrimTarget({
      rank: 1,
      label: "expired-previews",
      trim: async () => {
        order.push("previews");
      }
    });
    await governor.trim("moderate");
    expect(order).toEqual(["previews", "cold"]);
  });

  it("pressure listeners fire on level transitions", () => {
    const governor = new ResourceGovernor({ ...DEFAULT_BUDGETS, ioConcurrent: 1 });
    const listener = vi.fn();
    governor.onPressureChange(listener);
    const lease = governor.tryAcquire({ classes: { io: 1 }, label: "io" })!;
    expect(listener).toHaveBeenCalledWith(expect.objectContaining({ level: "critical" }));
    lease.release();
    expect(listener).toHaveBeenCalledTimes(2);
  });
});
