/**
 * Byte-budget LRU cache with admission (§100–§101). Caches are limited by
 * estimated bytes, never by object count.
 */

import type { CacheAdmissionPolicy, CacheKind } from "../../contracts/scheduler.js";

export interface ByteBudgetCacheStats {
  hits: number;
  misses: number;
  evictions: number;
  rejectedAdmissions: number;
  bytesHeld: number;
  size: number;
}

export class ByteBudgetCache<V> {
  private readonly map = new Map<string, { value: V }>();
  private bytesHeld = 0;
  readonly stats: ByteBudgetCacheStats = {
    hits: 0,
    misses: 0,
    evictions: 0,
    rejectedAdmissions: 0,
    bytesHeld: 0,
    size: 0
  };

  constructor(
    readonly kind: CacheKind,
    readonly byteBudget: number,
    private readonly sizeEstimate: (value: V) => number,
    private readonly policy: CacheAdmissionPolicy = { maxItemShareOfBudget: 0.25 }
  ) {}

  sizeEstimateOf(value: V): number {
    return this.sizeEstimate(value);
  }

  get(key: string): V | undefined {
    const entry = this.map.get(key);
    if (!entry) {
      this.stats.misses++;
      return undefined;
    }
    // LRU touch.
    this.map.delete(key);
    this.map.set(key, entry);
    this.stats.hits++;
    return entry.value;
  }

  /** Admission check (§100): huge one-shot payloads are not worth caching. */
  admit(estimatedBytes: number): boolean {
    if (estimatedBytes > this.byteBudget * this.policy.maxItemShareOfBudget) {
      this.stats.rejectedAdmissions++;
      return false;
    }
    return true;
  }

  set(key: string, value: V): void {
    const bytes = this.sizeEstimate(value);
    if (!this.admit(bytes)) return;
    const existing = this.map.get(key);
    if (existing) {
      this.bytesHeld -= this.sizeEstimate(existing.value);
      this.map.delete(key);
    }
    while (this.bytesHeld + bytes > this.byteBudget && this.map.size > 0) {
      const oldest = this.map.keys().next().value;
      if (oldest === undefined) break;
      const old = this.map.get(oldest)!;
      this.bytesHeld -= this.sizeEstimate(old.value);
      this.map.delete(oldest);
      this.stats.evictions++;
    }
    this.map.set(key, { value });
    this.bytesHeld += bytes;
    this.syncStats();
  }

  delete(key: string): void {
    const entry = this.map.get(key);
    if (!entry) return;
    this.bytesHeld -= this.sizeEstimate(entry.value);
    this.map.delete(key);
    this.syncStats();
  }

  clear(): void {
    this.map.clear();
    this.bytesHeld = 0;
    this.syncStats();
  }

  /** Evict everything (used by pressure ladder for the most recyclable layers). */
  trimAll(): number {
    const removed = this.map.size;
    this.clear();
    return removed;
  }

  entries(): IterableIterator<[string, V]> {
    return Array.from(this.map.entries(), ([k, v]) => [k, v.value] as [string, V])[Symbol.iterator]();
  }

  get currentBytes(): number {
    return this.bytesHeld;
  }

  get length(): number {
    return this.map.size;
  }

  private syncStats(): void {
    this.stats.bytesHeld = this.bytesHeld;
    this.stats.size = this.map.size;
  }
}
