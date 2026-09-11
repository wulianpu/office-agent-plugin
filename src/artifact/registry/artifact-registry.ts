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

const CONTEXT_BYTES_ESTIMATE = 8 * 1024; // metadata-level context footprint

interface BuildJob {
  key: string;
  consumers: Set<string>;
  abort: AbortController;
  promise: Promise<ArtifactContext>;
  cancelled: boolean;
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
    const fingerprint = await this.fingerprintFor(input.artifactRef);
    const key = contextKey(input.artifactRef, fingerprint, profile);

    const cached = this.completed.get(key);
    if (cached) {
      return this.leaseFor(cached, input.consumer, undefined, key);
    }

    let job = this.inFlight.get(key);
    if (!job) {
      const abort = new AbortController();
      const built: Promise<ArtifactContext> = runtime
        .createArtifactContext({
          ...input,
          profile,
          signal: abort.signal
        })
        .then((context) => {
          this.completed.set(key, context);
          this.inFlight.delete(key);
          return context;
        })
        .catch((error) => {
          this.inFlight.delete(key);
          throw error;
        });
      job = { key, consumers: new Set(), abort, promise: built, cancelled: false };
      this.inFlight.set(key, job);
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
    // Path resolution is injected by the service wiring to avoid a store dependency here.
    const path = this.pathResolver(ref);
    const fp = await fileFingerprint(path);
    return fingerprintKey(fp);
  }

  private pathResolver: (ref: ArtifactRef) => string = () => {
    throw new Error("ArtifactRegistry.setPathResolver not wired");
  };

  /** Service wiring: registry learns path resolution from the ArtifactStore. */
  setPathResolver(resolver: (ref: ArtifactRef) => string): void {
    this.pathResolver = resolver;
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
      // Consumer-aware cancellation (§24): no consumers left → cancel the build.
      job.cancelled = true;
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

export function contextKey(ref: ArtifactRef, fpKey: string, profile: ArtifactContextProfile): string {
  return `${ref}|${fpKey}|${profile === "full" ? "f" : "m"}`;
}
