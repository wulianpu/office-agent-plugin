/**
 * FastArtifactProbe (§79): hot-path probe for preview/open. Only stat + format
 * sniff + ZIP central directory. Never hashes, validates deeply, or renders
 * (PERF-07: strong hash must not block first preview).
 */

import type { ArtifactRef, OfficeFormat } from "../../contracts/ids.js";
import type { FastArtifactProbeResult } from "../../contracts/artifact.js";
import { fileFingerprint, fingerprintKey } from "../../support/fsx.js";
import { hasContentTypes, readZipIndex } from "../scanner/zip.js";

export interface ProbeOutput extends FastArtifactProbeResult {
  fingerprintKey: string;
}

export async function fastArtifactProbe(
  ref: ArtifactRef,
  path: string,
  format: OfficeFormat
): Promise<ProbeOutput> {
  const fingerprint = await fileFingerprint(path);
  let entryCount = 0;
  let packageSane = false;
  try {
    const index = await readZipIndex(path);
    entryCount = index.entries.length;
    packageSane = hasContentTypes(index);
  } catch {
    packageSane = false;
  }
  return {
    artifactRef: ref,
    format,
    fingerprint,
    entryCount,
    packageSane,
    probedAt: Date.now(),
    fingerprintKey: fingerprintKey(fingerprint)
  };
}
