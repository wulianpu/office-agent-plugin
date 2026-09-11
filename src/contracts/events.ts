/**
 * Phase 0 — Contracts: event plane (Design Doc §115–§119).
 */

import type { SessionEpoch, SessionId } from "./ids.js";

/** §117: durable session events (persisted, at-least-once, idempotent consumers). */
export interface SessionEvent<T = unknown> {
  eventId: string;
  sessionId: SessionId;
  sequence: bigint;
  sessionEpoch: SessionEpoch;
  type: SessionEventType;
  payload: T;
  createdAt: number;
}

export type SessionEventType =
  | "session.opened"
  | "session.closed"
  | "session.epoch-bumped"
  | "session.conflict"
  | "session.recovery-required"
  | "lease.acquired"
  | "lease.released"
  | "writer.fenced"
  | "candidate.created"
  | "candidate.state-changed"
  | "candidate.verified"
  | "candidate.verification-invalidated"
  | "candidate.rejected"
  | "revision.committed"
  | "commit.journal-updated"
  | "recovery.resolved"
  | "capability.changed";

/** §116: ephemeral editor signals — never persisted. */
export interface EditorSignal<T = unknown> {
  sessionId: SessionId;
  type: EditorSignalType;
  payload: T;
  at: number;
}

export type EditorSignalType =
  | "selectionChanged"
  | "viewportChanged"
  | "focusChanged"
  | "scroll"
  | "dirtyChanged"
  | "editorSuspended"
  | "editorResumed";

/** §118: at-least-once delivery with idempotent consumers; stale/duplicate events are ignored. */
export interface SessionEventStore {
  append(event: SessionEvent): Promise<void>;
  readSince(sessionId: SessionId, afterSequence: bigint, limit?: number): Promise<SessionEvent[]>;
}

export interface EventBus {
  publish(event: SessionEvent): void;
  subscribe(
    sessionId: SessionId | "*",
    listener: (event: SessionEvent) => void
  ): () => void;
}

export interface EditorSignalBus {
  publish(signal: EditorSignal): void;
  subscribe(
    sessionId: SessionId | "*",
    listener: (signal: EditorSignal) => void
  ): () => void;
}
