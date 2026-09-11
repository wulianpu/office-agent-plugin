/**
 * Phase 0 — Contracts: capability plane (Design Doc §57–§62, §122, §127).
 */

import type { CandidateId, DocumentId, RevisionId, SessionId } from "./ids.js";
import type { FencingToken } from "./ids.js";
import type { OfficeFormat } from "./ids.js";

/** §58: agents receive document capabilities, never real paths (INV-12). */
export interface DocumentCapability {
  documentId: DocumentId;
  sessionId: SessionId;
  permissions: {
    read: boolean;
    edit: boolean;
  };
  format: OfficeFormat;
}

/** §122: agent-referenced assets are opaque handles resolved by the runtime. */
export interface AssetCapability {
  assetId: string;
  /** Approved local artifact the runtime resolved for this asset. */
  artifactPath: string;
  mediaType: string;
  grantedAt: number;
}

/** §59: what an agent task operates on. */
export interface OfficeTaskContext {
  taskId: string;
  sessionId: SessionId;
  documentId: DocumentId;
  candidateId: CandidateId;
  baseRevisionId: RevisionId;
  fencingToken: FencingToken;
  mutationScope: MutationScope;
}

/** §60: scope fences agent mutations (INV tested in agent suite). */
export interface MutationScope {
  intent: string;
  allowedTargets?: string[];
  allowedParts?: string[];
  destructiveAllowed: boolean;
}

/** §61: mutating tool requests are idempotent commands. */
export interface MutationCommand<T = unknown> {
  commandId: string;
  idempotencyKey: string;
  candidateId: CandidateId;
  fencingToken: FencingToken;
  payload: T;
}

/** §62: audit receipt for each executed mutation. */
export interface MutationReceipt {
  receiptId: string;
  commandId: string;
  candidateId: CandidateId;
  engine: "officecli";
  affectedTargets: string[];
  completedAt: number;
}

export type OperationRisk = "low" | "medium" | "high" | "critical" | "denied";

export interface OperationPolicy {
  risk: OperationRisk;
  /** When risk exceeds the session threshold the mutation requires explicit approval. */
  requiresApproval: boolean;
  reason?: string;
}
