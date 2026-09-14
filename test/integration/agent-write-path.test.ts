/**
 * Agent write path integration (§158): the full Candidate workflow against a
 * real OfficeCLI engine — task → mutate → flush barrier → verify → handoff →
 * accept with atomic commit. Covers INV-01/04/05/06/07/10/11.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { sha256File } from "../../src/support/fsx.js";
import { newCommandId } from "../../src/support/ids.js";
import { openWorkspace } from "../helpers/fixtures.js";
import { createOfficeCliFixture } from "../helpers/officecli-fixture.js";

let ws: Awaited<ReturnType<typeof openWorkspace>>;
let fixture: Awaited<ReturnType<typeof createOfficeCliFixture>>;
/** CI without the OfficeCLI engine: the whole suite skips cleanly. */
const engineUp = await import("../../src/agent/officecli/officecli-adapter.js")
  .then(async (m) => {
    const probe = new m.OfficeCliAdapter({ timeoutMs: 5_000 }); // probe skips in seconds — never the 120s default
    return probe.version_().then(() => true).catch(() => false);
  })
  .catch(() => false);

let pptxPath: string;


afterAll(async () => {
  await ws?.cleanup().catch(() => undefined);
});

describe.skipIf(!engineUp)("Agent write path (§158)", () => {
  beforeAll(async () => {
    fixture = await createOfficeCliFixture();

    ws = await openWorkspace();
    pptxPath = await fixture.pptx(ws.root);
  });
  it("runs the full candidate workflow: mutate → flush → verify → accept", async () => {
    const ref = await ws.plugin.registerArtifact(pptxPath);
    const session = await ws.plugin.openSession(ref);
    // P0-6: open returns immediately (optimistic identity); writer paths
    // observe the strong revision via ensureStrongIdentity.
    const strong = await ws.plugin.service.sessions.ensureStrongIdentity(session.sessionId);
    const sourceHashBefore = await sha256File(pptxPath);
    expect(strong.contentHash).toBe(sourceHashBefore);

    const task = await ws.plugin.beginAgentTask(session.sessionId, {
      intent: "retitle slide 1",
      destructiveAllowed: false,
      allowedTargets: ["/slide[1]"]
    });
    expect(task.fencingToken).toBeGreaterThan(0n);

    // INV-01: the agent never touches the committed source.
    const candidate = ws.plugin.service.candidates.require(task.candidateId);
    const stagingPath = ws.plugin.service.store.resolvePath(candidate.artifactRef);
    expect(stagingPath).not.toBe(pptxPath);
    // INV-05: candidate bound to base revision + hash.
    expect(candidate.baseRevisionId).toBe(session.committedRevision.revisionId);
    expect(candidate.baseHash).toBe(sourceHashBefore);

    await ws.plugin.executeAgentMutation(task, {
      commandId: newCommandId(),
      idempotencyKey: "agent-full-1",
      payload: [
        { command: "set", path: "/slide[1]/shape[1]", props: { text: "Retitled by Agent" } }
      ]
    });

    // INV-06: resident mutation lives in memory; the candidate file on disk
    // still holds base bytes until the explicit flush barrier.
    const stagingBeforeFlush = await sha256File(stagingPath);
    expect(stagingBeforeFlush).toBe(sourceHashBefore);

    const flushedHash = await ws.plugin.flushAgentCandidate(task);
    expect(flushedHash).not.toBe(sourceHashBefore);
    // INV-01 continues: source untouched.
    expect(await sha256File(pptxPath)).toBe(sourceHashBefore);

    const report = await ws.plugin.verifyAgentCandidate(task, ["/slide[1]"]);
    expect(["structural", "engine", "visual"]).toContain(report.confidence);
    expect(report.contentHash).toBe(flushedHash);

    // INV-07: handoff closes the resident and releases the writer.
    await ws.plugin.finalizeAgentTask(task);
    expect(ws.plugin.service.residentPool.size()).toBe(0);
    expect(ws.plugin.service.leases.activeLease(session.sessionId)).toBeUndefined();

    const accepted = await ws.plugin.acceptCandidate(session.sessionId);
    expect(accepted.contextPromoted).toBe(true);
    const newHash = await sha256File(pptxPath);
    expect(newHash).toBe(flushedHash);

    const updated = ws.plugin.service.getSession(session.sessionId)!;
    expect(updated.committedRevision.sequence).toBe(2);
    expect(updated.committedRevision.origin).toBe("agent");
    await ws.plugin.closeSession(session.sessionId);
  });

  it("P1d: idempotency is candidate-scoped; same key + different payload conflicts (§61)", async () => {
    const ref = await ws.plugin.registerArtifact(pptxPath);
    const session = await ws.plugin.openSession(ref);

    // Task 1 uses the key with payload A.
    const task1 = await ws.plugin.beginAgentTask(session.sessionId, { intent: "t1", destructiveAllowed: false });
    await ws.plugin.executeAgentMutation(task1, {
      commandId: "cmd-a",
      idempotencyKey: "shared-key",
      payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Payload A" } }]
    });

    // Same key + DIFFERENT payload inside the same candidate → conflict.
    await expect(
      ws.plugin.executeAgentMutation(task1, {
        commandId: "cmd-b",
        idempotencyKey: "shared-key",
        payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Payload B" } }]
      })
    ).rejects.toMatchObject({ code: "idempotency-conflict" });

    await ws.plugin.finalizeAgentTask(task1);
    await ws.plugin.rejectCandidate(session.sessionId);

    // Task 2 (new candidate) MAY reuse the key — scope is per candidate.
    const task2 = await ws.plugin.beginAgentTask(session.sessionId, { intent: "t2", destructiveAllowed: false });
    await expect(
      ws.plugin.executeAgentMutation(task2, {
        commandId: "cmd-c",
        idempotencyKey: "shared-key",
        payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Payload C" } }]
      })
    ).resolves.toBeTruthy();
    await ws.plugin.finalizeAgentTask(task2);
    await ws.plugin.rejectCandidate(session.sessionId);
    await ws.plugin.closeSession(session.sessionId);
  });

  it("INV-04: replaying the same idempotencyKey returns the stored receipt without re-executing", async () => {
    const ref = await ws.plugin.registerArtifact(pptxPath);
    const session = await ws.plugin.openSession(ref);
    const task = await ws.plugin.beginAgentTask(session.sessionId, {
      intent: "idempotency check",
      destructiveAllowed: false
    });

    const command = {
      commandId: newCommandId(),
      idempotencyKey: "idem-1",
      payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Once only" } }]
    };
    const first = await ws.plugin.executeAgentMutation(task, command);
    const stagingPath = ws.plugin.service.store.resolvePath(
      ws.plugin.service.candidates.require(task.candidateId).artifactRef
    );
    // Flush so the disk bytes reflect exactly one mutation.
    await ws.plugin.flushAgentCandidate(task);
    const hashAfterFirst = await sha256File(stagingPath);

    const replay = await ws.plugin.executeAgentMutation(task, {
      ...command,
      commandId: newCommandId() // different command id, same idempotency key
    });
    expect(replay.receiptId).toBe(first.receiptId); // same stored receipt
    const hashAfterReplay = await sha256File(stagingPath);
    expect(hashAfterReplay).toBe(hashAfterFirst); // bytes unchanged — no re-execution

    // Cleanup: reject the candidate.
    await ws.plugin.finalizeAgentTask(task);
    await ws.plugin.rejectCandidate(session.sessionId);
    await ws.plugin.closeSession(session.sessionId);
  });
});

describe.skipIf(!engineUp)("Agent mutation concurrency & session teardown (issue #6)", () => {
  beforeAll(async () => {
    fixture = await createOfficeCliFixture();
    ws = await openWorkspace();
    pptxPath = await fixture.pptx(ws.root);
  });

  it("P0: two CONCURRENT same-key same-payload mutations execute once and return the same persisted receipt", async () => {
    const ref = await ws.plugin.registerArtifact(pptxPath);
    const session = await ws.plugin.openSession(ref);
    const task = await ws.plugin.beginAgentTask(session.sessionId, { intent: "concurrent retry", destructiveAllowed: false });
    const command = {
      commandId: "cmd-concurrent-a",
      idempotencyKey: "concurrent-retry-1",
      payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Once Only" } }]
    };
    const [first, second] = await Promise.all([
      ws.plugin.executeAgentMutation(task, command),
      // Same idempotencyKey + same payload, different commandId — a true
      // concurrent retry that previously double-executed.
      ws.plugin.executeAgentMutation(task, { ...command, commandId: "cmd-concurrent-b" })
    ]);
    expect(first.receiptId).toBe(second.receiptId); // one persisted receipt
    const executions = ws.plugin.service.repos.countIdempotencyExecutions(task.candidateId, "concurrent-retry-1");
    expect(executions).toBe(1); // exactly one durable row
    await ws.plugin.finalizeAgentTask(task);
    await ws.plugin.rejectCandidate(session.sessionId);
    await ws.plugin.closeSession(session.sessionId);
  });

  it("P0: concurrent same-key DIFFERENT-payload — one payload executes, the other conflicts with zero side effect", async () => {
    const ref = await ws.plugin.registerArtifact(pptxPath);
    const session = await ws.plugin.openSession(ref);
    const task = await ws.plugin.beginAgentTask(session.sessionId, { intent: "conflict race", destructiveAllowed: false });
    const results = await Promise.allSettled([
      ws.plugin.executeAgentMutation(task, {
        commandId: "cmd-race-1",
        idempotencyKey: "race-key",
        payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Payload One" } }]
      }),
      ws.plugin.executeAgentMutation(task, {
        commandId: "cmd-race-2",
        idempotencyKey: "race-key",
        payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Payload Two" } }]
      })
    ]);
    const fulfilled = results.filter((r) => r.status === "fulfilled");
    const rejected = results.filter((r) => r.status === "rejected");
    expect(fulfilled).toHaveLength(1);
    expect(rejected).toHaveLength(1);
    expect((rejected[0] as PromiseRejectedResult).reason).toMatchObject({ code: "idempotency-conflict" });
    expect(ws.plugin.service.repos.countIdempotencyExecutions(task.candidateId, "race-key")).toBe(1);
    await ws.plugin.finalizeAgentTask(task);
    await ws.plugin.rejectCandidate(session.sessionId);
    await ws.plugin.closeSession(session.sessionId);
  });

  it("P1-high: closeSession fully reclaims an ACTIVE agent task — resident, lease (owner=agent), candidate", async () => {
    const ref = await ws.plugin.registerArtifact(pptxPath);
    const session = await ws.plugin.openSession(ref);
    const task = await ws.plugin.beginAgentTask(session.sessionId, { intent: "teardown", destructiveAllowed: false });
    await ws.plugin.executeAgentMutation(task, {
      commandId: "cmd-teardown",
      idempotencyKey: "teardown-1",
      payload: [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Will Be Abandoned" } }]
    });
    // The mutation opened the engine resident for the candidate.
    expect(ws.plugin.service.agent.residentCount()).toBe(1);
    expect(ws.plugin.service.agent.activeTaskCount()).toBe(1);

    await ws.plugin.closeSession(session.sessionId);

    expect(ws.plugin.service.agent.activeTaskCount()).toBe(0); // task reclaimed
    expect(ws.plugin.service.agent.residentCount()).toBe(0); // resident evicted
    expect(ws.plugin.service.leases.hasActiveLease(session.sessionId)).toBe(false);
    // The durable audit records the TRUE owner — never human.
    const events = ws.plugin.service.repos.readEventsSince(session.sessionId, 0n, 200);
    const released = [...events].reverse().find((e) => e.type === "lease.released");
    expect((released?.payload as { owner?: string })?.owner).toBe("agent");
    // The abandoned candidate is explicitly failed — observable, not ownerless.
    const candidate = ws.plugin.service.repos.getCandidate(task.candidateId);
    expect(candidate?.state).toBe("failed");
    expect(candidate?.failureReason).toContain("session closed");
  });
});

describe.skipIf(!engineUp || process.platform !== "win32")(
  "Windows 8.3 short-path hardening (round 8)",
  () => {
    it("short-alias TEMP and file paths through the adapter: no abort, one identity", async (ctx) => {
      // Acceptance (issue #2): the parent process TEMP/TMP set to an 8.3
      // alias must not abort the engine (libuv fs-event assertion), and a
      // file referenced once by short alias and once by long path must stay
      // ONE engine identity (no double resident / double writer).
      const { execFile } = await import("node:child_process");
      const { promisify } = await import("node:util");
      const { mkdtemp, rm } = await import("node:fs/promises");
      const { tmpdir } = await import("node:os");
      const { join } = await import("node:path");
      const run = promisify(execFile);

      // Create a directory whose name forces an 8.3 alias (>8 chars) and ask
      // Windows for its short form. Volumes with 8.3 disabled return the
      // long form — the scenario is untestable there, skip honestly.
      const longDir = await mkdtemp(join(tmpdir(), "shortpathprobe-longname-"));
      let shortDir: string;
      try {
        shortDir = (
          await run("powershell", [
            "-NoProfile",
            "-Command",
            `(New-Object -ComObject Scripting.FileSystemObject).GetFolder('${longDir.replace(/\\/g, "\\\\")}').ShortPath`
          ])
        ).stdout.trim();
      } catch {
        shortDir = longDir;
      }
      if (shortDir === longDir) {
        await rm(longDir, { recursive: true, force: true }).catch(() => undefined);
        ctx.skip(); // no 8.3 aliases on this volume
      }

      const { OfficeCliAdapter } = await import("../../src/agent/officecli/officecli-adapter.js");
      const previousTemp = process.env.TEMP;
      const previousTmp = process.env.TMP;
      process.env.TEMP = shortDir;
      process.env.TMP = shortDir;
      try {
        const adapter = new OfficeCliAdapter({ timeoutMs: 60_000 });
        const longPath = join(longDir, "probe.docx");
        const shortPath = join(shortDir, "probe.docx");

        // Create via the LONG path, then operate through the SHORT alias —
        // the adapter must canonicalize both to one engine identity.
        await adapter.run(["create", longPath, "--json"]).catch(() => undefined);
        await adapter.runBatchStandalone(shortPath, [
          { command: "add", parent: "/body", type: "paragraph", props: { text: "short path probe" } }
        ]);
        await adapter.save(shortPath);
        await adapter.close(longPath); // long alias must release the SAME resident

        const check = await adapter.get(shortPath, "/body/paragraph[1]");
        expect(JSON.stringify(check)).toContain("short path probe");
        await adapter.close(shortPath).catch(() => undefined);
      } finally {
        process.env.TEMP = previousTemp;
        process.env.TMP = previousTmp;
        await rm(longDir, { recursive: true, force: true }).catch(() => undefined);
      }
    }, 120_000);

    it("short BASENAME alias of an existing file: one resident identity across aliases (round 9)", async (ctx) => {
      // Round 9 (issue #3): the JS realpath never expanded 8.3 names at all;
      // longFormPath now resolves the WHOLE path natively, so a short
      // BASENAME alias (PRESEN~1.DOCX) and the long name are ONE identity.
      const { execFile } = await import("node:child_process");
      const { promisify } = await import("node:util");
      const { mkdtemp, rm } = await import("node:fs/promises");
      const { tmpdir } = await import("node:os");
      const { join } = await import("node:path");
      const run = promisify(execFile);

      const dir = await mkdtemp(join(tmpdir(), "basename-alias-"));
      const longName = "Presentation Q3 2026.docx";
      const longPath = join(dir, longName);
      try {
        const { OfficeCliAdapter } = await import("../../src/agent/officecli/officecli-adapter.js");
        const adapter = new OfficeCliAdapter({ timeoutMs: 60_000 });
        await adapter.run(["create", longPath, "--json"]).catch(() => undefined);
        await adapter.runBatchStandalone(longPath, [
          { command: "add", parent: "/body", type: "paragraph", props: { text: "alias probe" } }
        ]);

        let shortPath: string;
        try {
          shortPath = (
            await run("powershell", [
              "-NoProfile",
              "-Command",
              `(New-Object -ComObject Scripting.FileSystemObject).GetFile('${longPath.replace(/\\/g, "\\\\")}').ShortPath`
            ])
          ).stdout.trim();
        } catch {
          shortPath = longPath;
        }
        // No 8.3 aliases on this volume (or name already 8.3-fit) — skip.
        if (shortPath === longPath) ctx.skip();
        expect(shortPath).not.toBe(longPath); // a REAL basename alias exists

        // Operate through the short alias — the adapter canonicalizes both
        // forms to one engine identity.
        const viaShort = await adapter.get(shortPath, "/body/paragraph[1]");
        expect(JSON.stringify(viaShort)).toContain("alias probe");
        // Closing via EITHER alias releases the SAME resident.
        await adapter.close(shortPath).catch(() => undefined);
        await adapter.close(longPath).catch(() => undefined);
        const again = await adapter.get(longPath, "/body/paragraph[1]");
        expect(JSON.stringify(again)).toContain("alias probe");
        await adapter.close(longPath).catch(() => undefined);
      } finally {
        await rm(dir, { recursive: true, force: true }).catch(() => undefined);
      }
    }, 120_000);
  }
);
