/**
 * Issue #7 (P1-high): untrusted OOXML/ZIP resource boundaries. ZIP metadata
 * is attacker-controlled — declared sizes are never trusted as memory
 * limits; DEFLATE output is bounded by ACTUAL bytes on both the buffered
 * and streaming paths; EOCD/ZIP64 numbers are validated before allocation.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { deflateRawSync } from "node:zlib";
import { buildZip } from "../helpers/zip-builder.js";
import {
  DEFAULT_ZIP_BUDGET,
  readZipEntry,
  readZipIndex,
  streamZipEntry,
  ZipFormatError
} from "../../src/artifact/scanner/zip.js";

let dir: string;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "zip-budget-"));
});

afterAll(async () => {
  await rm(dir, { recursive: true, force: true }).catch(() => undefined);
});

/** Hand-rolled ZIP writer with full control over the central directory —
 *  lets us lie about sizes the way an attacker would. */
function craftZip(parts: { name: string; method: number; data: Buffer; declaredSize?: number }[]): Buffer {
  const chunks: Buffer[] = [];
  const central: Buffer[] = [];
  let offset = 0;
  for (const part of parts) {
    const nameBuf = Buffer.from(part.name, "utf8");
    const lfh = Buffer.alloc(30);
    lfh.writeUInt32LE(0x04034b50, 0);
    lfh.writeUInt16LE(part.method, 8);
    lfh.writeUInt32LE(0, 14); // crc (not validated here)
    lfh.writeUInt32LE(part.data.length, 18); // compressed
    lfh.writeUInt32LE(part.declaredSize ?? part.data.length, 22); // uncompressed (can lie)
    lfh.writeUInt16LE(nameBuf.length, 26);
    chunks.push(lfh, nameBuf, part.data);
    const cde = Buffer.alloc(46);
    cde.writeUInt32LE(0x02014b50, 0);
    cde.writeUInt16LE(part.method, 10);
    cde.writeUInt32LE(part.data.length, 20); // compressed
    cde.writeUInt32LE(part.declaredSize ?? part.data.length, 24); // uncompressed (lie)
    cde.writeUInt16LE(nameBuf.length, 28);
    cde.writeUInt32LE(offset, 42);
    central.push(cde, nameBuf);
    offset += 30 + nameBuf.length + part.data.length;
  }
  const cd = Buffer.concat(central);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(parts.length, 8);
  eocd.writeUInt16LE(parts.length, 10);
  eocd.writeUInt32LE(cd.length, 12);
  eocd.writeUInt32LE(offset, 16);
  return Buffer.concat([...chunks, cd, eocd]);
}

describe("untrusted ZIP resource boundaries (issue #7)", () => {
  it("buffered inflate: declared-small entry with an exploding stream fails at the actual cap", async () => {
    // 4 MiB of zeros compresses tiny; the central directory DECLARES 100
    // bytes — the old code trusted the declaration and materialized ~4 MiB.
    const bomb = Buffer.alloc(4 * 1024 * 1024, 0x41);
    const compressed = deflateRawSync(bomb);
    const zip = craftZip([{ name: "evil.xml", method: 8, data: compressed, declaredSize: 100 }]);
    const path = join(dir, "bomb.docx");
    await writeFile(path, zip);
    const index = await readZipIndex(path);
    await expect(readZipEntry(path, index, "evil.xml", 64 * 1024)).rejects.toThrow(ZipFormatError);
  });

  it("streaming: cumulative decompressed output is capped by ACTUAL bytes; inflater destroyed, no further chunks", async () => {
    const bomb = Buffer.alloc(8 * 1024 * 1024, 0x42);
    const compressed = deflateRawSync(bomb);
    const zip = craftZip([{ name: "big.xml", method: 8, data: compressed, declaredSize: 50 }]);
    const path = join(dir, "stream-bomb.docx");
    await writeFile(path, zip);
    const index = await readZipIndex(path);
    const tinyBudget = { ...DEFAULT_ZIP_BUDGET, maxUncompressedEntryBytes: 64 * 1024 };
    let received = 0;
    await expect(
      (async () => {
        for await (const chunk of streamZipEntry(path, index, "big.xml", 64 * 1024, {
          budget: tinyBudget
        })) {
          received += chunk.length;
        }
      })()
    ).rejects.toThrow(ZipFormatError);
    expect(received).toBeLessThanOrEqual(tinyBudget.maxUncompressedEntryBytes + 64 * 1024);
  });

  it("oversized declared central directory fails BEFORE any large allocation", async () => {
    // entryCount beyond the budget cap — rejected pre-allocation.
    const zip = craftZip([{ name: "a.xml", method: 0, data: Buffer.from("x") }]);
    const path = join(dir, "many-entries.docx");
    await writeFile(path, zip);
    const tiny = { ...DEFAULT_ZIP_BUDGET, maxEntries: 0 };
    await expect(readZipIndex(path, tiny)).rejects.toThrow(/entry count exceeds budget/);
    const tinyCd = { ...DEFAULT_ZIP_BUDGET, maxCentralDirectoryBytes: 4 };
    await expect(readZipIndex(path, tinyCd)).rejects.toThrow(/central directory exceeds budget/);
  });

  it("stored entries: declared size << compressedSize fails BEFORE the big allocation (issue #7 reopen)", async () => {
    // method=0 with lying metadata: size=1KiB, compressedSize=8MiB. The
    // caller's 64KiB cap must reject it before Buffer.alloc(8MiB) semantics
    // matter — and the metadata mismatch itself is grounds to refuse.
    const payload = Buffer.alloc(8 * 1024 * 1024, 0x43);
    const zip = craftZip([{ name: "stored-lie.xml", method: 0, data: payload, declaredSize: 1024 }]);
    const path = join(dir, "stored-lie.docx");
    await writeFile(path, zip);
    const index = await readZipIndex(path);
    await expect(readZipEntry(path, index, "stored-lie.xml", 64 * 1024)).rejects.toThrow(
      /stored entry metadata mismatch/
    );
  });

  it("stored entries: honest large payloads still respect the caller cap", async () => {
    const payload = Buffer.alloc(256 * 1024, 0x44); // honest: size == compressedSize
    const zip = craftZip([{ name: "stored-big.xml", method: 0, data: payload }]);
    const path = join(dir, "stored-big.docx");
    await writeFile(path, zip);
    const index = await readZipIndex(path);
    await expect(readZipEntry(path, index, "stored-big.xml", 64 * 1024)).rejects.toThrow(
      /entry too large|stored entry exceeds budget/
    );
    const ok = await readZipEntry(path, index, "stored-big.xml", 1024 * 1024);
    expect(ok.length).toBe(256 * 1024);
  });

  it("per-entry ZIP64 extras with unsafe values are rejected before allocation/seek (issue #7 reopen)", async () => {
    // Build a ZIP whose single entry carries a malicious 0x0001 extra with
    // size sentinel + a >MAX_SAFE_INTEGER value.
    const data = Buffer.from("zip64-lies");
    const nameBuf = Buffer.from("z64.xml", "utf8");
    const extra = Buffer.alloc(32); // header 0x0001, dataSize 24: size(8)+csize(8)+offset(8)
    extra.writeUInt16LE(0x0001, 0);
    extra.writeUInt16LE(24, 2);
    extra.writeBigUInt64LE(BigInt("0xffffffffffffffff"), 4); // size sentinel replaced with huge
    extra.writeBigUInt64LE(BigInt(10), 12); // compressedSize
    extra.writeBigUInt64LE(BigInt(0), 20); // offset
    const lfh = Buffer.alloc(30);
    lfh.writeUInt32LE(0x04034b50, 0);
    lfh.writeUInt16LE(0, 8);
    lfh.writeUInt32LE(data.length, 18);
    lfh.writeUInt32LE(data.length, 22);
    lfh.writeUInt16LE(nameBuf.length, 26);
    const cde = Buffer.alloc(46);
    cde.writeUInt32LE(0x02014b50, 0);
    cde.writeUInt32LE(0xffffffff, 24); // size sentinel -> ZIP64 extra
    cde.writeUInt32LE(data.length, 20);
    cde.writeUInt16LE(nameBuf.length, 28);
    cde.writeUInt16LE(extra.length, 30); // extraLen
    cde.writeUInt32LE(0, 42);
    const cd = Buffer.concat([cde, nameBuf, extra]);
    const eocd = Buffer.alloc(22);
    eocd.writeUInt32LE(0x06054b50, 0);
    eocd.writeUInt16LE(1, 8);
    eocd.writeUInt16LE(1, 10);
    eocd.writeUInt32LE(cd.length, 12);
    eocd.writeUInt32LE(30 + nameBuf.length + data.length, 16);
    const zip = Buffer.concat([lfh, nameBuf, data, cd, eocd]);
    const path = join(dir, "z64-lie.docx");
    await writeFile(path, zip);
    await expect(readZipIndex(path)).rejects.toThrow(/ZIP64 entry size/);
  });

  it("normal fixtures do not regress (buffered + streaming + ZIP64-safe)", async () => {
    const zip = buildZip([
      { name: "[Content_Types].xml", data: "<?xml version=\"1.0\"?><Types/>" },
      { name: "word/document.xml", data: "<?xml version=\"1.0\"?><w>hello zip budget</w>" }
    ]);
    const path = join(dir, "normal.docx");
    await writeFile(path, zip);
    const index = await readZipIndex(path);
    expect(index.entries).toHaveLength(2);
    const doc = await readZipEntry(path, index, "word/document.xml");
    expect(doc.toString()).toContain("hello zip budget");
    let streamed = 0;
    for await (const chunk of streamZipEntry(path, index, "word/document.xml")) {
      streamed += chunk.length;
    }
    expect(streamed).toBe(doc.length);
  });
});
