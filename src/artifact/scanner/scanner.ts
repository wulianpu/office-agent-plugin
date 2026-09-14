/**
 * ArtifactScanner (§78): one strong pass produces contentHash + package
 * manifest + integrity, shared by verification, package diff and revision
 * metadata. Results are cached by content hash — repeated consumers of the
 * same bytes never rescan.
 */

import { crc32 } from "node:zlib";
import type { ArtifactScanResult, PackageManifest } from "../../contracts/artifact.js";
import { sha256File } from "../../support/fsx.js";
import { manifestFromIndex, readZipEntry, ZipFormatError } from "./zip.js";
import { ByteBudgetCache } from "../cache/byte-budget-cache.js";

export interface ScanPersistence {
  saveScan(row: { contentHash: string; size: string; manifest: string; corruptEntries: string[]; scannedAt: number }): Promise<void>;
}

export interface ScanOptions {
  /** Verify per-entry CRC by decompression (strong path; default true). */
  deepIntegrity?: boolean;
  /** Skip entries larger than this when CRC-checking (they still get hashed). */
  integrityMaxEntryBytes?: number;
}

const DEFAULT_INTEGRITY_MAX_ENTRY = 256 * 1024 * 1024;

export class ArtifactScanner {
  readonly scans = new ByteBudgetCache<{ hash: string; result: ArtifactScanResult }>(
    "ArtifactCache",
    64 * 1024 * 1024,
    (v) => 4096 + v.result.packageManifest.entries.length * 96
  );

  constructor(private readonly persistence?: ScanPersistence) {}

  /** P1-high (#4): scan-cache resident bytes join the global ledger. */
  wireCacheAccounting(reporter: (delta: number) => void): void {
    this.scans.setBytesReporter(reporter);
  }

  async scan(path: string, options: ScanOptions = {}): Promise<ArtifactScanResult> {
    const contentHash = await sha256File(path);
    const cached = this.scans.get(contentHash);
    if (cached) return cached.result;

    const result = await this.scanFresh(path, contentHash, options);
    if (this.scans.admit(this.scans.sizeEstimateOf({ hash: contentHash, result }))) {
      this.scans.set(contentHash, { hash: contentHash, result });
    }
    await this.persistence?.saveScan({
      contentHash,
      size: result.size.toString(),
      manifest: JSON.stringify(result.packageManifest),
      corruptEntries: result.integrity.corruptEntries,
      scannedAt: result.scannedAt
    });
    return result;
  }

  /** Hash-only fast variant when a manifest is not needed. */
  async hashOnly(path: string): Promise<string> {
    return sha256File(path);
  }

  private async scanFresh(
    path: string,
    contentHash: string,
    options: ScanOptions
  ): Promise<ArtifactScanResult> {
    const { readZipIndex } = await import("./zip.js");
    const { stat } = await import("node:fs/promises");
    const size = BigInt((await stat(path)).size);
    const manifest: PackageManifest = { entries: [] };
    const corruptEntries: string[] = [];

    try {
      const index = await readZipIndex(path);
      manifest.entries = manifestFromIndex(index).entries;
      if (options.deepIntegrity !== false) {
        const cap = options.integrityMaxEntryBytes ?? DEFAULT_INTEGRITY_MAX_ENTRY;
        for (const entry of index.entries) {
          if (entry.size > cap) continue;
          try {
            const data = await readZipEntry(path, index, entry.name, cap);
            if ((crc32(data) >>> 0) !== (entry.crc32 >>> 0)) {
              corruptEntries.push(entry.name);
            }
          } catch (error) {
            if (error instanceof ZipFormatError) corruptEntries.push(entry.name);
            else throw error;
          }
        }
      }
    } catch {
      // Not a readable ZIP: manifest stays empty; integrity records the failure.
      corruptEntries.push("__package__");
    }

    return {
      contentHash,
      size,
      packageManifest: manifest,
      integrity: { ok: corruptEntries.length === 0, corruptEntries },
      scannedAt: Date.now()
    };
  }

  dispose(): void {
    this.scans.clear();
  }
}
