/**
 * Issue #6 (P1-high): ResidentPool same-file opening single-flight and
 * exact ResourceLease ownership — concurrent acquire never double-opens,
 * never orphans a governor lease, and failures return the pool to baseline.
 */

import { describe, expect, it } from "vitest";
import { ResidentPool } from "../../src/agent/officecli/resident-pool.js";
import { ResourceGovernor, DEFAULT_BUDGETS } from "../../src/runtime/resources/resource-governor.js";
import type { OfficeCliAdapter } from "../../src/agent/officecli/officecli-adapter.js";

function fakeAdapter(options: { failFor?: string } = {}) {
  const state = { opens: 0, closes: 0 };
  const adapter = {
    async open(file: string): Promise<void> {
      state.opens++;
      if (options.failFor === file) throw new Error("engine open failed");
    },
    async close(): Promise<void> {
      state.closes++;
    }
  } as unknown as OfficeCliAdapter;
  return { adapter, state };
}

function smallBudgetPool(adapter: OfficeCliAdapter) {
  // nativeProcess budget of 1: any leaked resident permit would make the
  // NEXT acquire block/fail — a behavioral leak detector.
  const governor = new ResourceGovernor({
    ...DEFAULT_BUDGETS,
    nativeProcessConcurrent: 1,
    xlsxSidecarConcurrent: 1
  });
  return { governor, pool: new ResidentPool(adapter, governor) };
}

describe("ResidentPool single-flight (issue #6)", () => {
  it("10 concurrent acquire(sameFile): one open, one resident, one permit", async () => {
    const { adapter, state } = fakeAdapter();
    const { pool } = smallBudgetPool(adapter);
    const handles = await Promise.all(Array.from({ length: 10 }, () => pool.acquire("C:/tmp/book.docx")));
    expect(state.opens).toBe(1); // single-flight: one engine open
    expect(pool.size()).toBe(1);
    expect(handles).toHaveLength(10);
    await pool.dispose();
    expect(state.closes).toBe(1);
  });

  it("evict releases exactly one lease — a later different-file acquire succeeds under budget 1", async () => {
    const { adapter } = fakeAdapter();
    const { pool } = smallBudgetPool(adapter);
    await pool.acquire("C:/tmp/a.docx");
    expect(pool.size()).toBe(1);
    await pool.evict("C:/tmp/a.docx");
    expect(pool.size()).toBe(0);
    // Budget is 1: this only completes if the evicted lease was truly released.
    const guard = new Promise((resolve) => setTimeout(resolve, 3000));
    const next = pool.acquire("C:/tmp/b.docx");
    await Promise.race([next, guard.then(() => Promise.reject(new Error("permit leaked")))]).catch(
      (error) => {
        throw error;
      }
    );
    await next;
    await pool.dispose();
  });

  it("open failure: every joiner rejects, pool and permits return to baseline", async () => {
    const { adapter, state } = fakeAdapter({ failFor: "C:/tmp/bad.docx" });
    const { pool } = smallBudgetPool(adapter);
    const attempts = await Promise.allSettled([
      pool.acquire("C:/tmp/bad.docx"),
      pool.acquire("C:/tmp/bad.docx"),
      pool.acquire("C:/tmp/bad.docx")
    ]);
    expect(attempts.every((a) => a.status === "rejected")).toBe(true);
    expect(state.opens).toBe(1); // one attempt, shared by all joiners
    expect(pool.size()).toBe(0);
    // The failed open's permit was released exactly once — a subsequent
    // acquire proceeds under the budget of 1.
    await pool.acquire("C:/tmp/ok.docx");
    await pool.dispose();
  });
});
