/**
 * Phase 0 — Contracts: MCP tool plane (Design Doc §56–§57, §85).
 *
 * The tool surface stays small and never exposes filesystem/shell/process
 * management. Transport is stdio JSON-RPC (MCP); the tool core itself is
 * transport-agnostic so tests drive it directly.
 */

import type { OfficeFormat } from "./ids.js";
import type { MutationScope, MutationReceipt } from "./capabilities.js";
import type { VerificationReport } from "./verification.js";

export type OfficeToolName =
  | "office.inspect"
  | "office.query"
  | "office.edit"
  | "office.render"
  | "office.verify"
  | "office.capabilities";

export interface OfficeToolContext {
  /** Capability issued by the runtime; revoked sessions reject tool calls. */
  documentId?: string;
  sessionId: string;
}

export interface OfficeInspectInput extends OfficeToolContext {
  path?: string;
}

export interface OfficeInspectResult {
  format: OfficeFormat;
  node: unknown;
  /** Logical address of the returned node for follow-up queries. */
  path: string;
}

export interface OfficeQueryInput extends OfficeToolContext {
  selector: string;
}

export interface OfficeQueryResult {
  matches: Array<{ path: string; summary: Record<string, unknown> }>;
  truncated: boolean;
}

/** One officecli-shaped batch item, policy-checked before dispatch (§61). */
export interface OfficeEditItem {
  command: string;
  path?: string;
  parent?: string;
  selector?: string;
  type?: string;
  props?: Record<string, string>;
  to?: string;
  after?: string;
  before?: string;
  path2?: string;
  part?: string;
  content?: string;
}

export interface OfficeEditInput extends OfficeToolContext {
  intent: string;
  items: OfficeEditItem[];
  scope?: Partial<MutationScope>;
  idempotencyKey: string;
}

export interface OfficeEditResult {
  taskContext: {
    taskId: string;
    candidateId: string;
    baseRevisionId: string;
    fencingToken: string;
  };
  receipts: MutationReceipt[];
  /** Non-fatal policy/scope notes surfaced to the agent. */
  notes: string[];
}

export interface OfficeRenderInput extends OfficeToolContext {
  scope?: string;
  mode?: "text" | "outline";
}

export interface OfficeRenderResult {
  mode: "text" | "outline";
  sections: Array<{ label: string; text: string }>;
}

export interface OfficeVerifyInput extends OfficeToolContext {
  candidateId?: string;
}

export interface OfficeVerifyResult {
  report: VerificationReport;
}

export interface OfficeCapabilitiesResult {
  capabilities: Array<{
    format: OfficeFormat;
    editor: CapabilityStatus;
    agent: CapabilityStatus;
    verification: CapabilityStatus;
  }>;
  hostAdapters: Array<{ id: "powerpoint" | "wps"; status: CapabilityStatus }>;
  offline: boolean;
}

export type CapabilityStatusValue =
  | "available"
  | "degraded"
  | "unavailable";

export interface CapabilityStatus {
  status: CapabilityStatusValue;
  engine: string;
  reason?: string;
}
