/**
 * ResourceGovernor (§105–§107): owns global budgets across resource classes,
 * tracks pressure, and drives the eviction ladder. Protected state — unsaved
 * human edits, candidates, commit journal — is never evicted (§107).
 */

import type {
  MemoryTrimLevel
} from "../../contracts/artifact.js";
import type {
  ResourceClass,
  ResourceGovernor as IResourceGovernor,
  ResourceLease,
  ResourcePressure,
  ResourcePressureSnapshot,
  ResourceRequest,
  ResourceTrimTarget
} from "../../contracts/scheduler.js";

export interface ResourceBudgets {
  memoryBytes: number;
  cpuConcurrent: number;
  ioConcurrent: number;
  renderConcurrent: number;
  nativeProcessConcurrent: number;
  diskCacheBytes: number;
}

export const DEFAULT_BUDGETS: ResourceBudgets = {
  memoryBytes: 512 * 1024 * 1024,
  cpuConcurrent: Math.max(2, Math.floor(navigatorHardwareConcurrency() / 2)),
  ioConcurrent: 4,
  renderConcurrent: 2,
  nativeProcessConcurrent: 4,
  diskCacheBytes: 2 * 1024 * 1024 * 1024
};

function navigatorHardwareConcurrency(): number {
  return Number(process.env.OFFICE_PLUGIN_CPU_BUDGET ?? 8);
}

const CONCURRENT_CLASSES: Partial<Record<ResourceClass, keyof ResourceBudgets>> = {
  cpu: "cpuConcurrent",
  io: "ioConcurrent",
  render: "renderConcurrent",
  "native-process": "nativeProcessConcurrent"
};

/** Eviction ladder order (§107): lower rank trimmed first. */
export const EVICTION_LADDER = [
  "expired-previews",
  "decoded-visuals",
  "prefetch",
  "idle-residents",
  "warm-editors",
  "cold-artifact-cache"
] as const;

export class ResourceGovernor implements IResourceGovernor {
  private readonly held = new Map<string, ResourceRequest>();
  private readonly counters = new Map<ResourceClass, { used: number; protected: number }>();
  private memoryUsed = 0;
  private memoryProtected = 0;
  private nextLeaseId = 1;
  private readonly trimTargets: Array<ResourceTrimTarget & { active: boolean }> = [];
  private readonly listeners = new Set<(snapshot: ResourcePressureSnapshot) => void>();
  private readonly releaseListeners = new Set<() => void>();
  private lastPressure: ResourcePressure = "normal";
  private trimming = false;

  constructor(readonly budgets: ResourceBudgets = { ...DEFAULT_BUDGETS }) {
    for (const cls of ["memory", "cpu", "io", "render", "native-process", "disk-cache"] as ResourceClass[]) {
      this.counters.set(cls, { used: 0, protected: 0 });
    }
  }

  async acquire(request: ResourceRequest): Promise<ResourceLease> {
    for (let attempt = 0; ; attempt++) {
      const lease = this.tryAcquire(request);
      if (lease) return lease;
      await this.applyPressureRelief();
      if (attempt > 200) {
        throw new Error(`resource governor: could not satisfy request '${request.label}'`);
      }
      await new Promise((resolve) => setTimeout(resolve, Math.min(50 * (attempt + 1), 500)));
    }
  }

  tryAcquire(request: ResourceRequest): ResourceLease | undefined {
    for (const [cls, amount] of Object.entries(request.classes)) {
      if (!amount) continue;
      const key = cls as ResourceClass;
      if (key === "memory") continue; // memory is accounting-only, admission via pressure
      const budgetKey = CONCURRENT_CLASSES[key];
      if (budgetKey) {
        const counter = this.counters.get(key)!;
        if (counter.used + amount > (this.budgets[budgetKey] as number)) return undefined;
      }
    }
    const leaseId = `gov-${this.nextLeaseId++}`;
    for (const [cls, amount] of Object.entries(request.classes)) {
      if (!amount) continue;
      const key = cls as ResourceClass;
      const counter = this.counters.get(key);
      if (counter) {
        const delta = key === "memory" ? 0 : amount;
        counter.used += delta;
        if (request.protected) counter.protected += delta;
      }
    }
    const bytes = request.bytes ?? 0;
    this.memoryUsed += bytes;
    if (request.protected) this.memoryProtected += bytes;
    this.held.set(leaseId, request);
    this.refreshPressure();
    const self = this;
    return {
      leaseId,
      request,
      release(): void {
        self.releaseById(leaseId);
      }
    };
  }

  /** Fired on every release (scheduler re-pumps resource-blocked jobs). */
  private notifyRelease(): void {
    for (const listener of this.releaseListeners) listener();
  }

  private releaseById(leaseId: string): void {
    const request = this.held.get(leaseId);
    if (!request) return;
    this.held.delete(leaseId);
    for (const [cls, amount] of Object.entries(request.classes)) {
      if (!amount) continue;
      const key = cls as ResourceClass;
      const counter = this.counters.get(key);
      if (counter) {
        const delta = key === "memory" ? 0 : amount;
        counter.used = Math.max(0, counter.used - delta);
        if (request.protected) counter.protected = Math.max(0, counter.protected - delta);
      }
    }
    const bytes = request.bytes ?? 0;
    this.memoryUsed = Math.max(0, this.memoryUsed - bytes);
    if (request.protected) this.memoryProtected = Math.max(0, this.memoryProtected - bytes);
    this.refreshPressure();
    this.notifyRelease();
  }

  getPressure(): ResourcePressureSnapshot {
    return {
      level: this.lastPressure,
      utilization: this.utilization(),
      timestamp: Date.now()
    };
  }

  registerTrimTarget(target: ResourceTrimTarget): void {
    this.trimTargets.push({ ...target, active: true });
  }

  onPressureChange(listener: (snapshot: ResourcePressureSnapshot) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  /** P1-high-C: release notifications regardless of level transitions. */
  onRelease(listener: () => void): () => void {
    this.releaseListeners.add(listener);
    return () => this.releaseListeners.delete(listener);
  }

  utilization(): Partial<Record<ResourceClass, number>> {
    return {
      memory: this.memoryUsed / this.budgets.memoryBytes,
      cpu: this.counters.get("cpu")!.used / this.budgets.cpuConcurrent,
      io: this.counters.get("io")!.used / this.budgets.ioConcurrent,
      render: this.counters.get("render")!.used / this.budgets.renderConcurrent,
      "native-process":
        this.counters.get("native-process")!.used / this.budgets.nativeProcessConcurrent,
      "disk-cache": 0
    };
  }

  /** Bytes currently accounted (for telemetry/tests). */
  memoryUsage(): { used: number; protected: number; budget: number } {
    return { used: this.memoryUsed, protected: this.memoryProtected, budget: this.budgets.memoryBytes };
  }

  private refreshPressure(): void {
    const util = Object.values(this.utilization()).filter((v) => v !== undefined) as number[];
    const max = Math.max(0, ...util);
    const level: ResourcePressure =
      max >= 0.95 ? "critical" : max >= 0.8 ? "high" : max >= 0.6 ? "elevated" : "normal";
    if (level !== this.lastPressure) {
      this.lastPressure = level;
      const snapshot = this.getPressure();
      for (const listener of this.listeners) listener(snapshot);
    }
  }

  private async applyPressureRelief(): Promise<void> {
    if (this.trimming) return;
    this.trimming = true;
    try {
      const sorted = [...this.trimTargets].sort((a, b) => a.rank - b.rank);
      const level: MemoryTrimLevel = this.lastPressure === "critical" ? "aggressive" : "moderate";
      for (const target of sorted) {
        if (this.lastPressure === "normal") break;
        await target.trim(level).catch(() => undefined);
      }
    } finally {
      this.trimming = false;
    }
  }

  /** Manual trim entry point used by tab/residency policies (§95). */
  async trim(level: MemoryTrimLevel): Promise<void> {
    for (const target of [...this.trimTargets].sort((a, b) => a.rank - b.rank)) {
      await target.trim(level).catch(() => undefined);
    }
  }
}
