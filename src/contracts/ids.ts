/**
 * Phase 0 — Contracts: shared identifiers and primitives (Design Doc §15–§20, §154).
 *
 * Physical Office bytes are never exposed to callers by path; they are addressed
 * through opaque ArtifactRef handles owned by the ArtifactStore (§15).
 */

/** Opaque handle to a physical Office artifact. Real paths are known only to the ArtifactStore. */
export type ArtifactRef = string;

/** Opaque handle to a logical document identity (stable across revisions). */
export type DocumentId = string;

/** Opaque handle to a DocumentSession (§32). */
export type SessionId = string;

/** Opaque handle to a CommittedRevision (§36). */
export type RevisionId = string;

/** Opaque handle to a CandidateRevision (§37). */
export type CandidateId = string;

/** Opaque handle to a WriterLease (§38). */
export type LeaseId = string;

/** Monotonic token fencing stale writers (§39). */
export type FencingToken = bigint;

/** Session incarnation counter; bumped when a session is recovered/rebound (§32). */
export type SessionEpoch = number;

export type OfficeFormat = "docx" | "xlsx" | "pptx";

export type DocumentBackend =
  | "managed-file"
  | "external-powerpoint"
  | "external-wps";

export const OFFICE_FORMATS: readonly OfficeFormat[] = ["docx", "xlsx", "pptx"];

export function isOfficeFormat(value: string): value is OfficeFormat {
  return (OFFICE_FORMATS as readonly string[]).includes(value);
}

export function formatFromPath(path: string): OfficeFormat | undefined {
  const dot = path.lastIndexOf(".");
  if (dot < 0) return undefined;
  const ext = path.slice(dot + 1).toLowerCase();
  return isOfficeFormat(ext) ? ext : undefined;
}
