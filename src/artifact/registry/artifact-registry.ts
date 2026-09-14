/**
 * ArtifactRegistry (§21–§27): the only path to an ArtifactContext.
 *
 * - Completed contexts are cached and reference counted (ArtifactLease).
 * - Concurrent acquires join one in-flight build job (1 build → N consumers).
 * - Releasing the last consumer of an in-flight build cancels it; consumers
 *   of a completed context never block writers (MVCC, P4).
 */

import type {
  ArtifactBuildInput,
  ArtifactContext,
  ArtifactContextProfile,
  ArtifactLease,
  ArtifactRegistry as IArtifactRegistry,
  ArtifactVersionKey,
  FormatRuntime,
  MemoryTrimLevel
} from "../../contracts/artifact.js";
import type { ArtifactRef, OfficeFormat } from "../../contracts/ids.js";
import { fileFingerprint, fingerprintKey } from "../../support/fsx.js";
import { newId } from "../../support/ids.js";
import { ByteBudgetCache } from "../cache/byte-budget-cache.js";
import type { Scheduler } from "../../runtime/scheduler/scheduler.js";
import type { SchedulerPriorityName } from "../../contracts/scheduler.js";

const CONTEXT_BYTES_ESTIMATE = 8 * 1024; // metadata-level context footprint

interface BuildJob {
  key: string;
  consumers: Set<string>;
  abort: AbortController;
  promise: Promise<ArtifactContext>;
  cancelled: boolean;
  /** Round 10: scheduler handle for the build — cancel dequeues, promote
   *  implements priority inheritance for later joiners. */
  buildHandle?: { cancel(): void; promote(priority: SchedulerPriorityName): void };
}

export class ArtifactRegistry implements IArtifactRegistry {
  /** format → profile → runtime (P0-7: metadata/full build depths). */
  private readonly runtimes = new Map<OfficeFormat, Partial<Record<ArtifactContextProfile, FormatRuntime>>>();
  private readonly completed = new ByteBudgetCache<ArtifactContext>(
    "ArtifactCache",
    256 * 1024 * 1024,
    (context) => context.estimatedResidentBytes ?? CONTEXT_BYTES_ESTIMATE
  );
  private readonly inFlight = new Map<string, BuildJob>();
  private readonly leaseConsumers = new Map<string, string>();
  /** leaseId → { artifactRef, contextKey } for honest holders() introspection. */
  private readonly leaseIndex = new Map<string, { artifactRef: ArtifactRef; contextKey: string }>();
  private readonly inflightConsumersByLease = new Map<string, { key: string; consumer: string }>();
  /** Round 10: heavy builds (full-profile GenOffice parses) enter the shared
   *  Scheduler priority ladder + ResourceGovernor admission instead of
   *  running unregulated; 1 build → N consumers dedup is preserved. */
  private buildScheduler?: Scheduler;

  /** Service wiring: registry learns the shared scheduler for build admission. */
  setBuildScheduler(scheduler: Scheduler): void {
    this.buildScheduler = scheduler;
  }

  registerRuntime(runtime: FormatRuntime, profile: ArtifactContextProfile = "metadata"): void {
    const byProfile = this.runtimes.get(runtime.format) ?? {};
    byProfile[profile] = runtime;
    this.runtimes.set(runtime.format, byProfile);
  }

  runtimesSnapshot(): OfficeFormat[] {
    return [...this.runtimes.keys()];
  }

  hasRuntime(format: OfficeFormat, profile: ArtifactContextProfile): boolean {
    return this.runtimes.get(format)?.[profile] !== undefined;
  }

  async acquire(input: ArtifactBuildInput): Promise<ArtifactLease> {
    const profile = input.profile ?? "metadata";
    const runtime = this.runtimes.get(input.format)?.[profile];
    if (!runtime) {
      throw new Error(`no ${profile} FormatRuntime registered for ${input.format}`);
    }
    // Single stat feeds BOTH the context key and the transient peak
    // estimate — a second await here widened the join race window (#4).
    const fpInfo = await this.fingerprintInfo(input.artifactRef);
    const key = contextKey(input.artifactRef, fpInfo.key, profile);

    const cached = this.completed.get(key);
    if (cached) {
      return this.leaseFor(cached, input.consumer, undefined, key);
    }

    let job = this.inFlight.get(key);
    if (!job) {
      const abort = new AbortController();
      // P1-high (round 7): settle handlers are IDENTITY-GUARDED. After a
      // cancelled build is removed, a new job may already occupy this key —
      // an unguarded delete here would evict the NEW in-flight job (ABA),
      // and a cancelled resolve must never poison the completed cache.
      const self = this;
      const runBuild = () =>
        runtime.createArtifactContext({
          ...input,
          profile,
          signal: abort.signal
        });
      let built: Promise<ArtifactContext>;
      let buildHandle: BuildJob["buildHandle"];
      if (this.buildScheduler) {
        // Round 10 (P1-high): heavy context builds are scheduler/governor
        // admitted — a background full-profile parse no longer bypasses the
        // priority ladder. Full parses claim io+cpu; metadata stays light.
        // P1 (#4): full builds also reserve their TRANSIENT peak working set
        // (source bytes + parser multiplier) via estimatedBytes — the
        // governor's memory admission now bounds concurrent parse peaks;
        // the reservation is released when the job settles, after which the
        // cached context is accounted as RESIDENT bytes via the completed
        // cache's delta reporter (no double counting: phases are disjoint).
        const peakBytes = profile === "full" ? peakWorkingSet(fpInfo.size) : 0;
        const handle = this.buildScheduler.submit({
          label: `artifact-build:${input.artifactRef}:${profile}`,
          priority: input.priority,
          resources:
            profile === "full"
              ? { io: 1, cpu: 1, estimatedBytes: peakBytes }
              : { io: 1 },
          run: async (signal) => {
            if (signal.aborted) throw new DOMException("cancelled", "AbortError");
            return await runBuild();
          }
        });
        buildHandle = handle;
        built = handle.promise;
      } else {
        built = runBuild();
      }
      const settled: Promise<ArtifactContext> = built
        .then(function landed(context) {
          const owner = self.inFlight.get(key);
          if (!job!.cancelled) self.completed.set(key, context);
          if (owner === job!) self.inFlight.delete(key);
          return context;
        })
        .catch((error: unknown) => {
          const owner = self.inFlight.get(key);
          if (owner === job!) self.inFlight.delete(key);
          throw error;
        });
      job = { key, consumers: new Set(), abort, promise: settled, cancelled: false, buildHandle };
      this.inFlight.set(key, job);
    } else if (job.buildHandle) {
      // Priority inheritance for joiners: the shared build moves UP to the
      // newcomer's priority (Scheduler.promote is upward-only).
      job.buildHandle.promote(input.priority);
    }

    const leaseId = newId("lease");
    job.consumers.add(leaseId);
    this.leaseConsumers.set(leaseId, input.consumer);
    this.inflightConsumersByLease.set(leaseId, { key, consumer: input.consumer });

    // §24: a consumer's own abort only detaches it; the shared build aborts
    // when (and only when) the last consumer detaches.
    let onConsumerAbort: (() => void) | undefined;
    try {
      const context = await new Promise<ArtifactContext>((resolve, reject) => {
        if (input.signal) {
          onConsumerAbort = () => {
            this.detachInFlightConsumer(leaseId);
            reject(new DOMException("consumer aborted", "AbortError"));
          };
          if (input.signal.aborted) {
            onConsumerAbort();
          } else {
            input.signal.addEventListener("abort", onConsumerAbort, { once: true });
          }
        }
        job.promise.then(resolve, reject);
      });
      this.cleanupConsumerSignal(input.signal, onConsumerAbort);
      this.inflightConsumersByLease.delete(leaseId);
      return this.leaseFor(context, input.consumer, leaseId, key);
    } catch (error) {
      this.cleanupConsumerSignal(input.signal, onConsumerAbort);
      if (!this.inflightConsumersByLease.has(leaseId)) {
        // Detached consumer (individual abort) or shared failure.
        throw error;
      }
      this.detachInFlightConsumer(leaseId);
      throw error;
    }
  }

  private cleanupConsumerSignal(signal: AbortSignal | undefined, handler: (() => void) | undefined): void {
    if (signal && handler) signal.removeEventListener("abort", handler);
  }

  peek(key: ArtifactVersionKey, profile: ArtifactContextProfile = "metadata"): ArtifactContext | undefined {
    return this.completed.get(contextKey(key.artifactRef, fingerprintKey(key.fingerprint), profile));
  }

  holders(artifactRef: ArtifactRef): ReadonlyArray<{ leaseId: string; consumer: string }> {
    const out: Array<{ leaseId: string; consumer: string }> = [];
    for (const [leaseId, consumer] of this.leaseConsumers) {
      if (this.leaseIndex.get(leaseId)?.artifactRef === artifactRef) {
        out.push({ leaseId, consumer });
      }
    }
    return out;
  }

  /** Count live leases (debug/introspection for §22 diagrams). */
  leaseCount(): number {
    return this.leaseConsumers.size;
  }

  /**
   * §71–§72 Context Promotion: when a candidate's bytes become the committed
   * source (same content, new logical role), the already-parsed context is
   * re-registered under the source's version key — Accept costs zero reparse.
   * The stored context keeps its parsed engine model (WeakMap identity) even
   * though its artifactRef field still names the staging origin.
   */
  promote(from: { artifactRef: ArtifactRef; fingerprintKey: string; profile: ArtifactContextProfile }, to: { artifactRef: ArtifactRef; fingerprintKey: string; profile?: ArtifactContextProfile }): boolean {
    const fromKey = contextKey(from.artifactRef, from.fingerprintKey, from.profile);
    const context = this.completed.get(fromKey);
    if (!context) return false;
    const toKey = contextKey(to.artifactRef, to.fingerprintKey, to.profile ?? from.profile);
    // P1-high: MOVE (rekey) the candidate entry under the source key — the
    // parsed engine model moves with it, billed ONCE under the new key.
    const rebound: ArtifactContext = {
      ...context,
      artifactRef: to.artifactRef,
      version: { ...context.version, artifactRef: to.artifactRef },
      lastAccessAt: Date.now()
    };
    this.completed.delete(fromKey);
    if (this.completed.admit(this.completed.sizeEstimateOf(rebound))) {
      this.completed.set(toKey, rebound);
    }
    return true;
  }

  inFlightCount(): number {
    return this.inFlight.size;
  }

  async trim(level: MemoryTrimLevel): Promise<void> {
    if (level === "light") return;
    if (this.leaseConsumers.size === 0) {
      this.completed.clear();
      return;
    }
    // With live readers, only drop contexts nobody currently references.
    for (const [key] of Array.from(this.completed.entries())) {
      if (!this.isKeyHeldByInFlight(key)) this.completed.delete(key);
    }
    for (const [format, byProfile] of this.runtimes) {
      for (const runtime of Object.values(byProfile)) {
        await runtime?.trimMemory(level).catch(() => void format);
      }
    }
  }

  /** P1-high-B: a cache key pinned by ANY live lease must not be evicted —
   *  dropping it forces a duplicate rebuild next to the still-held context. */
  private isKeyHeldByInFlight(key: string): boolean {
    for (const indexed of this.leaseIndex.values()) {
      if (indexed.contextKey === key) return true;
    }
    return false;
  }

  private async fingerprintFor(ref: ArtifactRef): Promise<string> {
    return (await this.fingerprintInfo(ref)).key;
  }

  /** One stat → { context-key material, source size } (#4: the transient
   *  peak reservation reuses the SAME fingerprint — no second stat, no
   *  widened join race). */
  private async fingerprintInfo(ref: ArtifactRef): Promise<{ key: string; size: bigint }> {
    // Path resolution is injected by the service wiring to avoid a store dependency here.
    const path = this.pathResolver(ref);
    const fp = await fileFingerprint(path);
    return { key: fingerprintKey(fp), size: fp.size };
  }

  private pathResolver: (ref: ArtifactRef) => string = () => {
    throw new Error("ArtifactRegistry.setPathResolver not wired");
  };

  /** Service wiring: registry learns path resolution from the ArtifactStore. */
  setPathResolver(resolver: (ref: ArtifactRef) => string): void {
    this.pathResolver = resolver;
  }

  /** P1-high (#4): completed-context cache resident bytes join the global
   *  governor memory ledger (local budget stays as a second-layer cap). */
  wireCacheAccounting(reporter: (delta: number) => void): void {
    this.completed.setBytesReporter(reporter);
  }

  private leaseFor(
    context: ArtifactContext,
    consumer: string,
    reuseLeaseId?: string,
    contextKeyOf?: string
  ): ArtifactLease {
    const leaseId = reuseLeaseId ?? newId("lease");
    this.leaseConsumers.set(leaseId, consumer);
    this.leaseIndex.set(leaseId, { artifactRef: context.artifactRef, contextKey: contextKeyOf ?? "" });
    const self = this;
    return {
      leaseId,
      context,
      consumer,
      release(): void {
        self.leaseConsumers.delete(leaseId);
        self.leaseIndex.delete(leaseId);
      }
    };
  }

  private detachInFlightConsumer(leaseId: string): void {
    const entry = this.inflightConsumersByLease.get(leaseId);
    if (!entry) return;
    this.inflightConsumersByLease.delete(leaseId);
    this.leaseConsumers.delete(leaseId);
    this.leaseIndex.delete(leaseId);
    const job = this.inFlight.get(entry.key);
    if (!job) return;
    job.consumers.delete(leaseId);
    if (job.consumers.size === 0 && !job.cancelled) {
      // Consumer-aware cancellation (§24): no consumers left → cancel the
      // build. A STILL-QUEUED build is dequeued outright (round 10); a
      // running build stops via its AbortSignal.
      job.cancelled = true;
      job.buildHandle?.cancel();
      job.abort.abort();
      this.inFlight.delete(entry.key);
    }
  }

  async dispose(): Promise<void> {
    for (const job of this.inFlight.values()) {
      job.abort.abort();
    }
    this.inFlight.clear();
    this.completed.clear();
    for (const byProfile of this.runtimes.values()) {
      for (const runtime of Object.values(byProfile)) {
        await runtime?.dispose().catch(() => undefined);
      }
    }
    this.runtimes.clear();
  }
}

/**
 * Conservative transient peak estimate for a full GenOffice parse (#4 P1):
 * source buffer + parser model + decode working set. Capped so pathological
 * sources degrade to a bounded reservation instead of bypassing admission.
 */
function peakWorkingSet(sourceBytes: bigint): number {
  const MULT = 4n;
  const CAP = 256n * 1024n * 1024n;
  const FLOOR = 16n * 1024n * 1024n;
  const estimate = sourceBytes * MULT;
  const clamped = estimate > CAP ? CAP : estimate < FLOOR ? FLOOR : estimate;
  return Number(clamped);
}

export function contextKey(ref: ArtifactRef, fpKey: string, profile: ArtifactContextProfile): string {
  return `${ref}|${fpKey}|${profile === "full" ? "f" : "m"}`;
}
