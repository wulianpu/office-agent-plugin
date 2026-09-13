/**
 * Round 9 (issue #3): the sidecar client must survive a LIVE-BUT-WEDGED
 * child — requests carry deadlines, stderr is drained (bounded ring), stdin
 * errors settle pending, and pending.size returns to 0 on every terminal
 * path. Deterministic via a fake child substituted for the real spawn.
 */

import { describe, expect, it } from "vitest";
import { EventEmitter } from "node:events";
import { PassThrough, Writable } from "node:stream";
import type { ChildProcessWithoutNullStreams } from "node:child_process";
import { XlsxSidecarClient } from "../../src/vendor/genoffice/xlsx-sidecar.js";

class FakeChild extends EventEmitter {
  readonly stdout = new PassThrough();
  readonly stderr = new PassThrough();
  readonly stdin: Writable;
  killed = false;
  written: string[] = [];

  constructor() {
    super();
    const self = this;
    this.stdin = new Writable({
      write(chunk: Buffer, _enc, callback) {
        self.written.push(chunk.toString("utf8"));
        callback();
      }
    });
  }

  kill(): boolean {
    this.killed = true;
    // Real processes die asynchronously — dispose's once("exit") must be
    // attached before the event, as with a real child.
    setImmediate(() => this.emit("exit"));
    return true;
  }

  asChild(): ChildProcessWithoutNullStreams {
    return this as unknown as ChildProcessWithoutNullStreams;
  }

  /** Reply to the last written request envelope. */
  reply(result: unknown, ok = true): void {
    const envelope = JSON.parse(this.written[this.written.length - 1]!) as { requestId: string };
    this.stdout.write(`${JSON.stringify({ requestId: envelope.requestId, ok, result })}\n`);
  }
}

function makeClient(options?: { requestTimeoutMs?: number; unhealthyCooldownMs?: number }): {
  client: XlsxSidecarClient;
  child: FakeChild;
} {
  const child = new FakeChild();
  class TestClient extends XlsxSidecarClient {
    protected override startChild(): ChildProcessWithoutNullStreams {
      return child.asChild();
    }
  }
  // exePath only gates the existsSync pre-check — startChild never spawns it.
  const client = new TestClient(process.execPath, options);
  return { client, child };
}

describe("XlsxSidecarClient liveness (round 9, issue #3)", () => {
  it("dispatches a matching stdout line and clears pending", async () => {
    const { client, child } = makeClient({ requestTimeoutMs: 5_000 });
    const promise = client.open("C:/tmp/book.xlsx");
    expect(client.pendingCount()).toBe(1);
    child.reply({ sessionId: "s1", sheets: [] });
    const opened = await promise;
    expect(opened.sessionId).toBe("s1");
    expect(client.pendingCount()).toBe(0);
    await client.dispose();
  });

  it("live-but-wedged child: request rejects at the deadline, child killed, cooldown latches", async () => {
    const { client, child } = makeClient({ requestTimeoutMs: 80, unhealthyCooldownMs: 60_000 });
    const promise = client.open("C:/tmp/book.xlsx");
    await expect(promise).rejects.toMatchObject({ message: expect.stringContaining("timeout") });
    expect(child.killed).toBe(true);
    expect(client.pendingCount()).toBe(0);
    // The killed child exits asynchronously — let the exit event land so the
    // cooldown latch is observable (a request racing the exit gets
    // "sidecar exited", which equally fails fast into the fallback).
    await new Promise((resolve) => setImmediate(resolve));
    await new Promise((resolve) => setImmediate(resolve));
    await expect(client.open("C:/tmp/book.xlsx")).rejects.toMatchObject({
      message: expect.stringContaining("cooling down")
    });
    await client.dispose();
  });

  it("cooldown expires: a fresh spawn attempt is allowed (no permanent latch)", async () => {
    const { client, child } = makeClient({ requestTimeoutMs: 80, unhealthyCooldownMs: 10 });
    await expect(client.open("C:/tmp/book.xlsx")).rejects.toMatchObject({
      message: expect.stringContaining("timeout")
    });
    await new Promise((resolve) => setTimeout(resolve, 25));
    const second = client.open("C:/tmp/book.xlsx");
    child.reply({ sessionId: "s2", sheets: [] });
    await expect(second).resolves.toMatchObject({ sessionId: "s2" });
    await client.dispose();
  });

  it("stderr flood is drained into a bounded ring — no unbounded growth", async () => {
    const { client, child } = makeClient({ requestTimeoutMs: 60_000 });
    const promise = client.open("C:/tmp/book.xlsx");
    // Far beyond a typical pipe capacity; the client must keep draining.
    for (let i = 0; i < 512; i++) {
      child.stderr.write("x".repeat(1024) + "\n");
    }
    await new Promise((resolve) => setImmediate(resolve));
    expect(client.recentStderr().length).toBeLessThanOrEqual(9 * 1024); // ring cap + one chunk
    child.reply({ sessionId: "s3", sheets: [] });
    await promise; // the request never wedged despite the flood
    await client.dispose();
  });

  it("child exit settles pending with no orphans", async () => {
    const { client, child } = makeClient({ requestTimeoutMs: 60_000 });
    const promise = client.open("C:/tmp/book.xlsx");
    child.emit("exit");
    await expect(promise).rejects.toMatchObject({ message: "sidecar exited" });
    expect(client.pendingCount()).toBe(0);
    await client.dispose();
  });

  it("dispose settles pending and clears timers", async () => {
    const { client } = makeClient({ requestTimeoutMs: 60_000 });
    const promise = client.open("C:/tmp/book.xlsx");
    // Attach the rejection handler BEFORE dispose — the rejection fires
    // synchronously inside dispose and must not sit unhandled across the
    // exit wait.
    const expectation = expect(promise).rejects.toMatchObject({ message: "sidecar disposed" });
    await client.dispose();
    await expectation;
    expect(client.pendingCount()).toBe(0);
  });
});
