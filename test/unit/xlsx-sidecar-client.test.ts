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
import { XlsxSidecarClient, sidecarPreviewWindow } from "../../src/vendor/genoffice/xlsx-sidecar.js";

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

describe("sidecarPreviewWindow scope boundaries (round 10 reopen #3)", () => {
  function fakeClient(sheets: Array<{ id: string; name: string; rowCount?: number; columnCount?: number }>, reads: Array<Record<string, number>> = [], cells: Array<{ row: number; column: number; value: string }> = []) {
    return {
      open: async () => ({ sessionId: "s", sheets }),
      readRange: async (_s: string, _id: string, range: Record<string, number>) => {
        reads.push(range);
        return { cells };
      },
      close: async () => undefined
    } as unknown as XlsxSidecarClient;
  }

  it("an explicit MISSING sheet fails closed: empty result, no reads, no other sheets", async () => {
    const reads: Array<Record<string, number>> = [];
    const fake = fakeClient([{ id: "1", name: "Data", rowCount: 20 }], reads);
    const out = await sidecarPreviewWindow(fake, "book.xlsx", { sheet: "DoesNotExist", range: { fromRow: 1, toRow: 3, fromCol: 1, toCol: 3 } });
    expect(out).toEqual([]); // never the real "Data" sheet under a missing-sheet key
    expect(reads).toHaveLength(0);
  });

  it("near-edge ranges clamp to the REMAINING extent, not the full sheet", async () => {
    // 20-row sheet. A18:C30 asks for 13 rows; only 3 remain from row 18.
    const reads: Array<Record<string, number>> = [];
    const fake = fakeClient([{ id: "1", name: "S", rowCount: 20, columnCount: 3 }], reads, [
      { row: 17, column: 0, value: "A18" },
      { row: 18, column: 0, value: "A19" },
      { row: 19, column: 0, value: "A20" }
    ]);
    const out = await sidecarPreviewWindow(fake, "book.xlsx", {
      range: { fromRow: 18, toRow: 30, fromCol: 1, toCol: 3 },
      maxRows: 13,
      maxCols: 3
    });
    // endRow clamped to the sheet edge (19), not startRow0 + 13 - 1 (29).
    expect(reads[0]).toEqual({ startRow: 17, endRow: 19, startColumn: 0, endColumn: 2 });
    expect(out[0]!.window).toHaveLength(3); // the legal remainder, never a fixed 5-row retry
    expect(out[0]!.window[0]![0]).toBe("A18");
  });

  it("a full in-extent range reads everything it asked for (A10:C20 of 20 rows -> 11 rows)", async () => {
    const reads: Array<Record<string, number>> = [];
    const cells = Array.from({ length: 11 }, (_, i) => ({ row: 9 + i, column: 0, value: "A" + (10 + i) }));
    const fake = fakeClient([{ id: "1", name: "S", rowCount: 20, columnCount: 3 }], reads, cells);
    const out = await sidecarPreviewWindow(fake, "book.xlsx", {
      range: { fromRow: 10, toRow: 20, fromCol: 1, toCol: 3 },
      maxRows: 11,
      maxCols: 3
    });
    expect(reads[0]).toEqual({ startRow: 9, endRow: 19, startColumn: 0, endColumn: 2 });
    expect(out[0]!.window).toHaveLength(11);
  });

  it("maxEntries can NARROW an explicit range but never EXTEND it (round 10 reopen #4)", async () => {
    // A5:A10 with a caller-passed window of 100 rows: the range span (6)
    // wins the min — endRow must stop at row 10, not run toward A104.
    const reads: Array<Record<string, number>> = [];
    const cells = Array.from({ length: 6 }, (_, i) => ({ row: 4 + i, column: 0, value: "A" + (5 + i) }));
    const fake = fakeClient([{ id: "1", name: "S", rowCount: 200, columnCount: 3 }], reads, cells);
    const out = await sidecarPreviewWindow(fake, "book.xlsx", {
      range: { fromRow: 5, toRow: 10, fromCol: 1, toCol: 1 },
      maxRows: 100,
      maxCols: 1
    });
    expect(reads[0]).toEqual({ startRow: 4, endRow: 9, startColumn: 0, endColumn: 0 });
    expect(out[0]!.window).toHaveLength(6); // rows 5..10 exactly
  });

  it("maxEntries narrowing wins: A5:A10 + 3-entry cap -> rows 5..7", async () => {
    const reads: Array<Record<string, number>> = [];
    const fake = fakeClient([{ id: "1", name: "S", rowCount: 200, columnCount: 3 }], reads, [
      { row: 4, column: 0, value: "A5" },
      { row: 5, column: 0, value: "A6" },
      { row: 6, column: 0, value: "A7" }
    ]);
    const out = await sidecarPreviewWindow(fake, "book.xlsx", {
      range: { fromRow: 5, toRow: 10, fromCol: 1, toCol: 1 },
      maxRows: 3,
      maxCols: 1
    });
    expect(reads[0]).toEqual({ startRow: 4, endRow: 6, startColumn: 0, endColumn: 0 });
    expect(out[0]!.window).toHaveLength(3);
  });

  it("an origin beyond the declared extent yields an explicitly empty window without reads", async () => {
    const reads: Array<Record<string, number>> = [];
    const fake = fakeClient([{ id: "1", name: "S", rowCount: 20, columnCount: 3 }], reads);
    const out = await sidecarPreviewWindow(fake, "book.xlsx", {
      range: { fromRow: 25, toRow: 30, fromCol: 1, toCol: 3 }
    });
    expect(out[0]!.window).toEqual([]);
    expect(reads).toHaveLength(0);
    const fakeCol = fakeClient([{ id: "1", name: "S", rowCount: 20, columnCount: 3 }], reads);
    const outCol = await sidecarPreviewWindow(fakeCol, "book.xlsx", {
      range: { fromRow: 1, toRow: 5, fromCol: 7, toCol: 9 }
    });
    expect(outCol[0]!.window).toEqual([]);
    expect(reads).toHaveLength(0);
  });
});

describe("sidecarPreviewWindow range origin (round 10 reopen)", () => {
  it("forwards the 1-based range as native 0-based read coordinates AND rebases absolute responses (round 10 reopen)", async () => {
    const ranges: Array<Record<string, number>> = [];
    const fake = {
      open: async () => ({
        sessionId: "s",
        sheets: [{ id: "sh1", name: "Data", rowCount: 500, columnCount: 60 }]
      }),
      // Native filters by range but returns ABSOLUTE worksheet coordinates.
      readRange: async (_s: string, _id: string, range: Record<string, number>) => {
        ranges.push(range);
        return {
          cells: [
            { row: 9, column: 3, value: "D10" },
            { row: 29, column: 5, value: "F30" }
          ]
        };
      },
      close: async () => undefined
    } as unknown as XlsxSidecarClient;
    const sheets = await sidecarPreviewWindow(fake, "book.xlsx", {
      sheet: "Data",
      range: { fromRow: 10, toRow: 30, fromCol: 4, toCol: 6 },
      maxRows: 21,
      maxCols: 3
    });
    // D10 (1-based) is native (row 9, col 3); the window spans 21x3.
    expect(ranges[0]).toEqual({ startRow: 9, endRow: 29, startColumn: 3, endColumn: 5 });
    // Rebased to viewport-local: absolute (9,3) → window[0][0]; (29,5) → window[20][2].
    // Pre-fix these cells were DROPPED as col >= maxCols.
    expect(sheets[0]!.window[0]![0]).toBe("D10");
    expect(sheets[0]!.window[20]![2]).toBe("F30");
    expect(sheets[0]!.window).toHaveLength(21);
  });

  it("A1-origin windows keep their legacy mapping", async () => {
    const fake = {
      open: async () => ({ sessionId: "s", sheets: [{ id: "sh1", name: "S", rowCount: 50 }] }),
      readRange: async () => ({ cells: [{ row: 0, column: 1, value: "B1" }] }),
      close: async () => undefined
    } as unknown as XlsxSidecarClient;
    const sheets = await sidecarPreviewWindow(fake, "book.xlsx", { maxRows: 5, maxCols: 4 });
    expect(sheets[0]!.window[0]![1]).toBe("B1");
  });

  it("default window still reads from the origin (0,0)", async () => {
    const ranges: Array<Record<string, number>> = [];
    const fake = {
      open: async () => ({ sessionId: "s", sheets: [{ id: "sh1", name: "S", rowCount: 50 }] }),
      readRange: async (_s: string, _id: string, range: Record<string, number>) => {
        ranges.push(range);
        return { cells: [] };
      },
      close: async () => undefined
    } as unknown as XlsxSidecarClient;
    await sidecarPreviewWindow(fake, "book.xlsx", { maxRows: 5, maxCols: 4 });
    expect(ranges[0]).toEqual({ startRow: 0, endRow: 4, startColumn: 0, endColumn: 3 });
  });
});

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
