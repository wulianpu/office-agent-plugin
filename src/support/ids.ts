import { randomBytes } from "node:crypto";

const ALPHABET = "0123456789abcdefghjkmnpqrstvwxyz";

function randomBase32(length: number): string {
  const bytes = randomBytes(length);
  let out = "";
  for (let i = 0; i < length; i++) {
    out += ALPHABET[bytes[i]! % 32];
  }
  return out;
}

/** Prefixed, sortable-enough opaque identifiers (artifact_, sess_, cand_, …). */
export function newId(prefix: string): string {
  return `${prefix}_${Date.now().toString(36)}${randomBase32(10)}`;
}

export function newArtifactRef(): string {
  return newId("art");
}
export function newSessionId(): string {
  return newId("sess");
}
export function newDocumentId(): string {
  return newId("doc");
}
export function newRevisionId(): string {
  return newId("rev");
}
export function newCandidateId(): string {
  return newId("cand");
}
export function newLeaseId(): string {
  return newId("lease");
}
export function newEventId(): string {
  return newId("evt");
}
export function newCommandId(): string {
  return newId("cmd");
}
export function newReceiptId(): string {
  return newId("rcpt");
}
export function newCommitId(): string {
  return newId("cmt");
}
export function newTaskId(): string {
  return newId("task");
}
export function newLeaseTokenId(): string {
  return newId("rls");
}
export function newRequestId(): string {
  return newId("req");
}
