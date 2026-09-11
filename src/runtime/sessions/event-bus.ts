/**
 * Event buses (§116–§118): durable SessionEvents are persisted with a
 * per-session monotonic sequence and delivered at-least-once; ephemeral
 * EditorSignals are never persisted.
 */

import type { SessionEpoch, SessionId } from "../../contracts/ids.js";
import type {
  EditorSignal,
  EditorSignalBus,
  EventBus,
  SessionEvent,
  SessionEventType
} from "../../contracts/events.js";
import { newEventId } from "../../support/ids.js";
import type { RuntimeRepositories } from "../persistence/repositories.js";

type SessionListener = (event: SessionEvent) => void;
type SignalListener = (signal: EditorSignal) => void;

export class DurableEventBus implements EventBus {
  private readonly listeners = new Map<SessionId | "*", Set<SessionListener>>();
  private readonly inflight: Array<Promise<void>> = [];

  constructor(private readonly repos: RuntimeRepositories) {}

  /** Persist + distribute. Append failures propagate (durable events must land). */
  publish(event: SessionEvent): void {
    this.repos.appendEvent(event);
    for (const listener of this.listeners.get(event.sessionId) ?? []) {
      this.inflight.push(Promise.resolve().then(() => listener(event)).catch(() => undefined));
    }
    for (const listener of this.listeners.get("*") ?? []) {
      this.inflight.push(Promise.resolve().then(() => listener(event)).catch(() => undefined));
    }
  }

  async emit(
    sessionId: SessionId,
    sessionEpoch: SessionEpoch,
    type: SessionEventType,
    payload: unknown
  ): Promise<SessionEvent> {
    const event: SessionEvent = {
      eventId: newEventId(),
      sessionId,
      sequence: this.repos.lastEventSequence(sessionId) + 1n,
      sessionEpoch,
      type,
      payload,
      createdAt: Date.now()
    };
    this.publish(event);
    return event;
  }

  subscribe(sessionId: SessionId | "*", listener: SessionListener): () => void {
    let set = this.listeners.get(sessionId);
    if (!set) {
      set = new Set();
      this.listeners.set(sessionId, set);
    }
    set.add(listener);
    return () => set!.delete(listener);
  }

  readSince(sessionId: SessionId, afterSequence: bigint, limit?: number): Promise<SessionEvent[]> {
    return Promise.resolve(this.repos.readEventsSince(sessionId, afterSequence, limit));
  }

  async drain(): Promise<void> {
    await Promise.all(this.inflight.splice(0));
  }
}

export class EditorSignalBusImpl implements EditorSignalBus {
  private readonly listeners = new Map<SessionId | "*", Set<SignalListener>>();

  publish(signal: EditorSignal): void {
    for (const listener of this.listeners.get(signal.sessionId) ?? []) listener(signal);
    for (const listener of this.listeners.get("*") ?? []) listener(signal);
  }

  subscribe(sessionId: SessionId | "*", listener: SignalListener): () => void {
    let set = this.listeners.get(sessionId);
    if (!set) {
      set = new Set();
      this.listeners.set(sessionId, set);
    }
    set.add(listener);
    return () => set!.delete(listener);
  }
}
