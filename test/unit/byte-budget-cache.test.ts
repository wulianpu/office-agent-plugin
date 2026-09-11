import { describe, expect, it } from "vitest";
import { ByteBudgetCache } from "../../src/artifact/cache/byte-budget-cache.js";

describe("ByteBudgetCache (§100–§101)", () => {
  it("evicts LRU by byte budget, never by object count", () => {
    const cache = new ByteBudgetCache<string>("ArtifactCache", 1000, (v) => v.length, {
      maxItemShareOfBudget: 0.5
    });
    cache.set("a", "x".repeat(400));
    cache.set("b", "x".repeat(400));
    cache.set("c", "x".repeat(400)); // 1200 > 1000 → evicts "a" (LRU)
    expect(cache.get("a")).toBeUndefined();
    expect(cache.get("b")).toBeDefined();
    expect(cache.get("c")).toBeDefined();
    expect(cache.currentBytes).toBe(800);
    expect(cache.stats.evictions).toBe(1);
  });

  it("admission rejects oversized one-shot items", () => {
    const cache = new ByteBudgetCache<string>("DecodedAssetCache", 100, (v) => v.length, {
      maxItemShareOfBudget: 0.25
    });
    cache.set("big", "x".repeat(30)); // 30% of budget > 25% → rejected
    expect(cache.get("big")).toBeUndefined();
    expect(cache.stats.rejectedAdmissions).toBe(1);
  });

  it("LRU touch protects recently used entries", () => {
    const cache = new ByteBudgetCache<string>("VisualPreviewCache", 800, (v) => v.length, {
      maxItemShareOfBudget: 0.5
    });
    cache.set("a", "x".repeat(400));
    cache.set("b", "x".repeat(400));
    cache.get("a"); // touch a → b becomes LRU
    cache.set("c", "x".repeat(400)); // evicts b
    expect(cache.get("a")).toBeDefined();
    expect(cache.get("b")).toBeUndefined();
    expect(cache.get("c")).toBeDefined();
  });
});
