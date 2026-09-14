/**
 * Minimal OOXML-grade ZIP reader: central directory index, ZIP64 support,
 * buffered and streaming entry access. Serves FastArtifactProbe (§79),
 * ArtifactScanner (§78) and the preview renderers without third-party deps.
 */

import { open } from "node:fs/promises";
import { createInflateRaw, inflateRawSync } from "node:zlib";
import type { PackageManifest, PackageManifestEntry } from "../../contracts/artifact.js";

const EOCD_SIG = 0x06054b50;
const ZIP64_EOCD_LOCATOR_SIG = 0x07064b50;
const ZIP64_EOCD_SIG = 0x06064b50;
const CDE_SIG = 0x02014b50;
const LFH_SIG = 0x04034b50;

export class ZipFormatError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ZipFormatError";
  }
}

/** P1-high (#7): untrusted-ZIP resource limits — one budget shared by the
 *  buffered and streaming read paths (and callers). ZIP metadata is
 *  attacker-controlled: nothing is allocated or trusted from it before these
 *  bounds hold, and DEFLATE output is capped by ACTUAL bytes, never by the
 *  declared uncompressed size. */
export interface ZipReadBudget {
  maxCentralDirectoryBytes: number;
  maxEntries: number;
  maxCompressedEntryBytes: number;
  maxUncompressedEntryBytes: number;
  /** Archive-wide cumulative decompressed budget for one streaming read. */
  maxTotalUncompressedBytes: number;
}

export const DEFAULT_ZIP_BUDGET: ZipReadBudget = {
  maxCentralDirectoryBytes: 16 * 1024 * 1024,
  maxEntries: 65_536,
  maxCompressedEntryBytes: 512 * 1024 * 1024,
  maxUncompressedEntryBytes: 512 * 1024 * 1024,
  maxTotalUncompressedBytes: 512 * 1024 * 1024
};

function safeNonNegativeInt(value: number, what: string): number {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new ZipFormatError(`malformed ${what}: ${String(value)}`);
  }
  return value;
}

export interface ZipEntry extends PackageManifestEntry {
  localHeaderOffset: number;
}

export interface ZipIndex {
  entries: ZipEntry[];
  entryByName: Map<string, ZipEntry>;
}

export async function readZipIndex(
  path: string,
  budget: ZipReadBudget = DEFAULT_ZIP_BUDGET
): Promise<ZipIndex> {
  const handle = await open(path, "r");
  try {
    const fileSize = (await handle.stat()).size;
    if (fileSize < 22) throw new ZipFormatError("file too small to be a ZIP");

    // Locate EOCD: scan the tail (comment can be up to 64KB).
    const tailLen = Math.min(fileSize, 22 + 65_536);
    const tail = Buffer.alloc(tailLen);
    await handle.read(tail, 0, tailLen, BigInt(fileSize - tailLen));
    let eocdOffsetInTail = -1;
    for (let i = tail.length - 22; i >= 0; i--) {
      if (tail.readUInt32LE(i) === EOCD_SIG) {
        eocdOffsetInTail = i;
        break;
      }
    }
    if (eocdOffsetInTail < 0) throw new ZipFormatError("EOCD not found");
    const eocd = tail.subarray(eocdOffsetInTail, eocdOffsetInTail + 22);
    let entryCount = eocd.readUInt16LE(10);
    let cdSize = Number(eocd.readUInt32LE(12));
    let cdOffset = Number(eocd.readUInt32LE(16));

    // ZIP64: sentinel values mean real numbers live in the ZIP64 EOCD record.
    if (entryCount === 0xffff || cdOffset === 0xffffffff || cdSize === 0xffffffff) {
      const locatorPos = eocdOffsetInTail - 20;
      if (locatorPos < 0 || tail.readUInt32LE(locatorPos) !== ZIP64_EOCD_LOCATOR_SIG) {
        throw new ZipFormatError("ZIP64 sentinels without locator");
      }
      const z64Offset = safeNonNegativeInt(
        Number(tail.readBigUInt64LE(locatorPos + 8)),
        "ZIP64 EOCD offset"
      );
      const z64 = Buffer.alloc(56);
      await handle.read(z64, 0, 56, BigInt(z64Offset));
      if (z64.readUInt32LE(0) !== ZIP64_EOCD_SIG) throw new ZipFormatError("ZIP64 EOCD not found");
      entryCount = safeNonNegativeInt(Number(z64.readBigUInt64LE(32)), "ZIP64 entry count");
      cdSize = safeNonNegativeInt(Number(z64.readBigUInt64LE(40)), "ZIP64 central-directory size");
      cdOffset = safeNonNegativeInt(
        Number(z64.readBigUInt64LE(48)),
        "ZIP64 central-directory offset"
      );
    }

    // P1-high (#7): attacker-controlled metadata gets hard bounds BEFORE any
    // allocation — a huge cdSize/entryCount cannot force a giant Buffer.alloc
    // or an unbounded parse loop even when it technically fits the file.
    if (cdSize > budget.maxCentralDirectoryBytes) {
      throw new ZipFormatError(
        `central directory exceeds budget: ${cdSize} > ${budget.maxCentralDirectoryBytes}`
      );
    }
    if (entryCount > budget.maxEntries) {
      throw new ZipFormatError(`entry count exceeds budget: ${entryCount} > ${budget.maxEntries}`);
    }
    if (cdOffset + cdSize > fileSize) throw new ZipFormatError("central directory out of range");

    const cd = Buffer.alloc(cdSize);
    await handle.read(cd, 0, cdSize, BigInt(cdOffset));

    const entries: ZipEntry[] = [];
    let pos = 0;
    for (let i = 0; i < entryCount; i++) {
      if (pos + 46 > cd.length || cd.readUInt32LE(pos) !== CDE_SIG) {
        throw new ZipFormatError(`central directory entry ${i} malformed`);
      }
      const method = cd.readUInt16LE(pos + 10);
      const crc32 = cd.readUInt32LE(pos + 16);
      let compressedSize = cd.readUInt32LE(pos + 20);
      let size = cd.readUInt32LE(pos + 24);
      const nameLen = cd.readUInt16LE(pos + 28);
      const extraLen = cd.readUInt16LE(pos + 30);
      const commentLen = cd.readUInt16LE(pos + 32);
      let localHeaderOffset = cd.readUInt32LE(pos + 42);
      const name = cd.subarray(pos + 46, pos + 46 + nameLen).toString("utf8");

      // ZIP64 extra field (0x0001) carries real sizes/offset when sentinels present.
      let extraPos = pos + 46 + nameLen;
      const extraEnd = extraPos + extraLen;
      while (extraPos + 4 <= extraEnd) {
        const headerId = cd.readUInt16LE(extraPos);
        const dataSize = cd.readUInt16LE(extraPos + 2);
        if (headerId === 0x0001) {
          let f = extraPos + 4;
          if (size === 0xffffffff && f + 8 <= extraPos + 4 + dataSize) {
            size = Number(cd.readBigUInt64LE(f));
            f += 8;
          }
          if (compressedSize === 0xffffffff && f + 8 <= extraPos + 4 + dataSize) {
            compressedSize = Number(cd.readBigUInt64LE(f));
            f += 8;
          }
          if (localHeaderOffset === 0xffffffff && f + 8 <= extraPos + 4 + dataSize) {
            localHeaderOffset = Number(cd.readBigUInt64LE(f));
          }
          break;
        }
        extraPos += 4 + dataSize;
      }

      entries.push({ name, compressedSize, size, crc32, method, localHeaderOffset });
      pos += 46 + nameLen + extraLen + commentLen;
    }

    return { entries, entryByName: new Map(entries.map((e) => [e.name, e])) };
  } finally {
    await handle.close();
  }
}

export function manifestFromIndex(index: ZipIndex): PackageManifest {
  return {
    entries: index.entries.map(({ name, compressedSize, size, crc32, method }) => ({
      name,
      compressedSize,
      size,
      crc32,
      method
    }))
  };
}

/** Read one entry fully into memory (fine for metadata/small parts). */
export async function readZipEntry(
  path: string,
  index: ZipIndex,
  name: string,
  maxBytes = 64 * 1024 * 1024,
  budget: ZipReadBudget = DEFAULT_ZIP_BUDGET
): Promise<Buffer> {
  const entry = index.entryByName.get(name);
  if (!entry) throw new ZipFormatError(`entry not found: ${name}`);
  const handle = await open(path, "r");
  try {
    const lfh = Buffer.alloc(30);
    await handle.read(lfh, 0, 30, BigInt(entry.localHeaderOffset));
    if (lfh.readUInt32LE(0) !== LFH_SIG) throw new ZipFormatError(`local header missing: ${name}`);
    const nameLen = lfh.readUInt16LE(26);
    const extraLen = lfh.readUInt16LE(28);
    const dataOffset = entry.localHeaderOffset + 30 + nameLen + extraLen;

    if (entry.size > maxBytes) {
      throw new ZipFormatError(`entry too large for buffered read: ${name} (${entry.size})`);
    }
    if (entry.compressedSize > budget.maxCompressedEntryBytes) {
      throw new ZipFormatError(
        `compressed entry exceeds budget: ${name} (${entry.compressedSize})`
      );
    }
    const raw = Buffer.alloc(entry.compressedSize);
    await handle.read(raw, 0, entry.compressedSize, BigInt(dataOffset));
    if (entry.method === 0) return raw;
    if (entry.method === 8) {
      // P1-high (#7): the DECLARED size is metadata — actual inflate output
      // is capped by zlib's maxOutputLength and verified afterwards. A small
      // declared size with an exploding stream fails closed.
      try {
        const out = inflateRawSync(raw, { maxOutputLength: maxBytes });
        if (out.length !== entry.size) {
          throw new ZipFormatError(
            `inflated entry size mismatch: ${name} (declared ${entry.size}, actual ${out.length})`
          );
        }
        return out;
      } catch (error) {
        if (error instanceof ZipFormatError) throw error;
        const message = String((error as Error)?.message ?? error);
        throw new ZipFormatError(`inflate failed or exceeded budget for ${name}: ${message}`);
      }
    }
    throw new ZipFormatError(`unsupported compression method ${entry.method}: ${name}`);
  } finally {
    await handle.close();
  }
}

/**
 * Stream one entry as chunks. Keeps large sheet XML out of the JS heap
 * (Working Set principle §31).
 */
export async function* streamZipEntry(
  path: string,
  index: ZipIndex,
  name: string,
  chunkSize = 4 * 1024 * 1024,
  options: { budget?: ZipReadBudget; signal?: AbortSignal } = {}
): AsyncGenerator<Buffer> {
  const entry = index.entryByName.get(name);
  if (!entry) throw new ZipFormatError(`entry not found: ${name}`);
  const budget = options.budget ?? DEFAULT_ZIP_BUDGET;
  const handle = await open(path, "r");
  let inflater: ReturnType<typeof createInflateRaw> | undefined;
  try {
    const lfh = Buffer.alloc(30);
    await handle.read(lfh, 0, 30, BigInt(entry.localHeaderOffset));
    if (lfh.readUInt32LE(0) !== LFH_SIG) throw new ZipFormatError(`local header missing: ${name}`);
    const nameLen = lfh.readUInt16LE(26);
    const extraLen = lfh.readUInt16LE(28);
    const dataOffset = entry.localHeaderOffset + 30 + nameLen + extraLen;

    if (entry.method === 0) {
      if (entry.compressedSize > budget.maxUncompressedEntryBytes) {
        throw new ZipFormatError(
          `stored entry exceeds budget: ${name} (${entry.compressedSize})`
        );
      }
      let position = dataOffset;
      let remaining = entry.compressedSize;
      while (remaining > 0) {
        if (options.signal?.aborted) throw new ZipFormatError("stream aborted");
        const read = Math.min(chunkSize, remaining);
        const buf = Buffer.alloc(read);
        await handle.read(buf, 0, read, BigInt(position));
        yield buf;
        position += read;
        remaining -= read;
      }
      return;
    }
    if (entry.method !== 8) throw new ZipFormatError(`unsupported compression method ${entry.method}`);

    // Queue + backpressure: compressed writes respect zlib's drain signal.
    // P1-high (#7): cumulative DECOMPRESSED output is budgeted by ACTUAL
    // emitted bytes (never the declared size) — exceeding the per-entry or
    // archive-wide cap destroys the inflater immediately and the consumer
    // receives no further chunks.
    const queue: Buffer[] = [];
    let failure: unknown;
    let emitted = 0;
    let overBudget = false;
    inflater = createInflateRaw();
    inflater.on("data", (chunk: Buffer) => {
      emitted += chunk.length;
      if (
        emitted > budget.maxUncompressedEntryBytes ||
        emitted > budget.maxTotalUncompressedBytes
      ) {
        overBudget = true;
        inflater!.destroy();
        return;
      }
      queue.push(chunk);
    });
    inflater.on("error", (error: unknown) => {
      failure = error;
    });

    let position = dataOffset;
    let remaining = entry.compressedSize;
    while (remaining > 0) {
      if (options.signal?.aborted) throw new ZipFormatError("stream aborted");
      const read = Math.min(chunkSize, remaining);
      const buf = Buffer.alloc(read);
      await handle.read(buf, 0, read, BigInt(position));
      position += read;
      remaining -= read;
      if (overBudget) break;
      if (!inflater.write(buf)) {
        // A destroyed inflater never drains — race the wait against the
        // budget teardown so over-budget destroys cannot hang the loop.
        await new Promise<void>((resolve) => {
          inflater!.once("drain", resolve);
          inflater!.once("close", resolve);
          if (overBudget || failure) resolve();
        });
        if (overBudget || failure) break;
      }
      while (queue.length > 0) {
        if (failure) throw failure;
        yield queue.shift()!;
      }
      if (failure) throw failure;
    }
    if (overBudget) {
      throw new ZipFormatError(
        `streamed entry exceeded decompressed budget: ${name} (>${budget.maxUncompressedEntryBytes} actual bytes; declared ${entry.size})`
      );
    }
    await new Promise<void>((resolve, reject) => {
      inflater!.once("end", resolve);
      inflater!.once("close", resolve);
      inflater!.once("error", reject);
      inflater!.end();
    });
    // The budget teardown may land only after end() — flush the final state.
    if (overBudget) {
      throw new ZipFormatError(
        `streamed entry exceeded decompressed budget: ${name} (>${budget.maxUncompressedEntryBytes} actual bytes; declared ${entry.size})`
      );
    }
    while (queue.length > 0) {
      if (failure) throw failure;
      yield queue.shift()!;
    }
    if (failure) throw failure;
  } finally {
    inflater?.destroy();
    await handle.close();
  }
}

export function hasContentTypes(index: ZipIndex): boolean {
  return index.entryByName.has("[Content_Types].xml");
}
