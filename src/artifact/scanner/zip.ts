/**
 * Minimal OOXML-grade ZIP reader: central directory index, ZIP64 support,
 * buffered and streaming entry access. Serves FastArtifactProbe (§79),
 * ArtifactScanner (§78) and the preview renderers without third-party deps.
 */

import { open } from "node:fs/promises";
import { createInflateRaw } from "node:zlib";
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

export interface ZipEntry extends PackageManifestEntry {
  localHeaderOffset: number;
}

export interface ZipIndex {
  entries: ZipEntry[];
  entryByName: Map<string, ZipEntry>;
}

export async function readZipIndex(path: string): Promise<ZipIndex> {
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
      const z64Offset = Number(tail.readBigUInt64LE(locatorPos + 8));
      const z64 = Buffer.alloc(56);
      await handle.read(z64, 0, 56, BigInt(z64Offset));
      if (z64.readUInt32LE(0) !== ZIP64_EOCD_SIG) throw new ZipFormatError("ZIP64 EOCD not found");
      entryCount = Number(z64.readBigUInt64LE(32));
      cdSize = Number(z64.readBigUInt64LE(40));
      cdOffset = Number(z64.readBigUInt64LE(48));
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
  maxBytes = 64 * 1024 * 1024
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
    const raw = Buffer.alloc(entry.compressedSize);
    await handle.read(raw, 0, entry.compressedSize, BigInt(dataOffset));
    if (entry.method === 0) return raw;
    if (entry.method === 8) {
      const { inflateRawSync } = await import("node:zlib");
      return inflateRawSync(raw);
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
  chunkSize = 4 * 1024 * 1024
): AsyncGenerator<Buffer> {
  const entry = index.entryByName.get(name);
  if (!entry) throw new ZipFormatError(`entry not found: ${name}`);
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
      let position = dataOffset;
      let remaining = entry.compressedSize;
      while (remaining > 0) {
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
    const queue: Buffer[] = [];
    let failure: unknown;
    inflater = createInflateRaw();
    inflater.on("data", (chunk: Buffer) => {
      queue.push(chunk);
    });
    inflater.on("error", (error: unknown) => {
      failure = error;
    });

    let position = dataOffset;
    let remaining = entry.compressedSize;
    while (remaining > 0) {
      const read = Math.min(chunkSize, remaining);
      const buf = Buffer.alloc(read);
      await handle.read(buf, 0, read, BigInt(position));
      position += read;
      remaining -= read;
      if (!inflater.write(buf)) {
        await new Promise<void>((resolve) => inflater!.once("drain", resolve));
      }
      while (queue.length > 0) {
        if (failure) throw failure;
        yield queue.shift()!;
      }
      if (failure) throw failure;
    }
    await new Promise<void>((resolve, reject) => {
      inflater!.once("end", resolve);
      inflater!.once("error", reject);
      inflater!.end();
    });
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
