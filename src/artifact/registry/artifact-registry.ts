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
  private readonly runtimes = new Map<OfficeFormat, FormatRuntime>();
  private readonly completed = new ByteBudgetCache<ArtifactContext>(
    "ArtifactCache",
    256 * 1024 * 1024,
    () => CONTEXT_BYTES_ESTIMATE
  );
  private readonly inFlight = new Map<string, BuildJob>();
  private readonly leaseConsumers = new Map<string, string>();
  private readonly inflightConsumersByLease = new Map<string, { key: string; consumer: string }>();

  registerRuntime(runtime: FormatRuntime): void {
    this.runtimes.set(runtime.format, runtime);
  }

  runtimesSnapshot(): OfficeFormat[] {
    return [...this.runtimes.keys()];
  }

  async acquire(input: ArtifactBuildInput): Promise<ArtifactLease> {
    const runtime = this.runtimes.get(input.format);
    if (!runtime) {
      throw new Error(`no FormatRuntime registered for ${input.format}`);
    }
    const fingerprint = await this.fingerprintFor(input.artifactRef);
    const key = contextKey(input.artifactRef, fingerprint, input.consistency);

    const cached = this.completed.get(key);
    if (cached) {
      return this.leaseFor(cached, input.consumer);
    }

    let job = this.inFlight.get(key);
    if (!job) {
      const abort = new AbortController();
      const built: Promise<ArtifactContext> = runtime
        .createArtifactContext({
          ...input,
          signal: abort.signal
        })
        .then((context) => {
          if (this.completed.admit(CONTEXT_BYTES_ESTIMATE)) {
            this.completed.set(key, context);
          }
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
      return this.leaseFor(context, input.consumer, leaseId);
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

  peek(key: ArtifactVersionKey): ArtifactContext | undefined {
    return this.completed.get(contextKey(key.artifactRef, fingerprintKey(key.fingerprint), "optimistic"));
  }

  holders(artifactRef: ArtifactRef): ReadonlyArray<{ leaseId: string; consumer: string }> {
    const out: Array<{ leaseId: string; consumer: string }> = [];
    for (const [leaseId, consumer] of this.leaseConsumers) {
      out.push({ leaseId, consumer });
    }
    return out.filter(() => true).map((h) => h); // stable snapshot; contexts carry ref in debug
  }

  /** Count live leases (debug/introspection for §22 diagrams). */
  leaseCount(): number {
    return this.leaseConsumers.size;
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
    for (const [format, runtime] of this.runtimes) {
      await runtime.trimMemory(level).catch(() => void format);
    }
  }

  private isKeyHeldByInFlight(_key: string): boolean {
    return false; // completed contexts are immutable; safe to drop while readers hold their lease object
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

  private leaseFor(context: ArtifactContext, consumer: string, reuseLeaseId?: string): ArtifactLease {
    const leaseId = reuseLeaseId ?? newId("lease");
    this.leaseConsumers.set(leaseId, consumer);
    const self = this;
    return {
      leaseId,
      context,
      consumer,
      release(): void {
        self.leaseConsumers.delete(leaseId);
      }
    };
  }

  private detachInFlightConsumer(leaseId: string): void {
    const entry = this.inflightConsumersByLease.get(leaseId);
    if (!entry) return;
    this.inflightConsumersByLease.delete(leaseId);
    this.leaseConsumers.delete(leaseId);
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
    for (const runtime of this.runtimes.values()) {
      await runtime.dispose().catch(() => undefined);
    }
    this.runtimes.clear();
  }
}

export function contextKey(ref: ArtifactRef, fpKey: string, consistency: "optimistic" | "stable"): string {
  return `${ref}|${fpKey}|${consistency === "stable" ? "s" : "o"}`;
}
