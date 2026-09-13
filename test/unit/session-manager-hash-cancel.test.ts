/**
 * P1 (round 7): closeSession must cancel the background strong-hash job in
 * BOTH states — running (AbortController → sha256File chunk loop) and QUEUED
 * (scheduler handle cancel → dequeue). A queued job that is only
 * signal-aborted stays in the queue until dispatch, occupying queue depth and
 * briefly acquiring IO resources after the session is gone.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { RuntimeDatabase } from "../../src/runtime/persistence/database.js";
import { RuntimeRepositories } from "../../src/runtime/persistence/repositories.js";
import { DurableEventBus } from "../../src/runtime/sessions/event-bus.js";
import { RevisionLog } from "../../src/runtime/revisions/revision-log.js";
import { Scheduler } from "../../src/runtime/scheduler/scheduler.js";
import { SessionManager } from "../../src/runtime/sessions/session-manager.js";
import type { ArtifactStore } from "../../src/artifact/store/artifact-store.js";
import type { ArtifactScanner } from "../../src/artifact/scanner/scanner.js";
import type { CandidateManager } from "../../src/runtime/candidates/candidate-manager.js";

let dir: string;
let db: RuntimeDatabase;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "hash-cancel-test-"));
  db = new RuntimeDatabase(join(dir, "rt.db"));
});

afterAll(async () => {
  db?.close();
  await rm(dir, { recursive: true, force: true }).catch(() => undefined);
});

describe("SessionManager background hash cancellation (P1)", () => {
  it("close() dequeues a QUEUED strong-hash job, not just its signal", async () => {
    const fixturePath = join(dir, "source.docx");
    await writeFile(fixturePath, Buffer.alloc(4096, 7));

    const store = {
      formatOf: () => "docx",
      resolvePath: () => fixturePath,
      tryResolvePath: () => undefined
    } as unknown as ArtifactStore;
    const repos = new RuntimeRepositories(db);
    const events = new DurableEventBus(repos);
    const revisions = new RevisionLog(repos);
    // maxConcurrent=1 with a never-resolving blocker → open()'s hash job
    // stays QUEUED, exactly the state the fix targets.
    const scheduler = new Scheduler({ maxConcurrent: 1 });
    const blocker = scheduler.submit({
      label: "blocker",
      priority: "INTERACTIVE",
      resources: { io: 1 },
      run: () => new Promise<never>(() => undefined)
    });
    void blocker.promise.catch(() => undefined);
    const manager = new SessionManager(
      store,
      {} as ArtifactScanner,
      repos,
      events,
      revisions,
      {} as CandidateManager,
      scheduler
    );

    const session = await manager.open("art_hash_cancel");
    expect(scheduler.queueDepth).toBe(1); // hash job is queued behind the blocker

    await manager.close(session.sessionId);

    // The queued job left the queue NOW — not "when the scheduler eventually
    // dispatches it". Pre-fix, only the AbortController fired and the job
    // stayed queued (queueDepth 1) until dispatch.
    expect(scheduler.queueDepth).toBe(0);
    expect(scheduler.stats.cancelled).toBe(1);

    blocker.cancel();
  });
});
