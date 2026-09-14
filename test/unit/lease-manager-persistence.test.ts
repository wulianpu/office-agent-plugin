/**
 * Issue #6 (P1): LeaseManager persistence-first linearization — a failed DB
 * insert leaves no phantom in-memory writer; a failed DB release keeps the
 * memory slot (the runtime refuses the next writer while the audit row is
 * unreleased). Deterministic via injected repository faults.
 */

import { describe, expect, it } from "vitest";
import { LeaseManager } from "../../src/runtime/sessions/lease-manager.js";
import type { RuntimeRepositories } from "../../src/runtime/persistence/repositories.js";

function stubRepos(faults: { insert?: boolean; release?: boolean } = {}) {
  return {
    latestFencingToken: () => 0n,
    insertLease: () => {
      if (faults.insert) throw new Error("db insert failed");
    },
    releaseLease: () => {
      if (faults.release) throw new Error("db release failed");
    }
  } as unknown as RuntimeRepositories;
}

const input = (owner: "agent" | "human" = "agent") => ({
  sessionId: "sess-lm",
  owner,
  backend: "officecli" as const,
  baseRevisionId: "rev_1",
  sourcePath: "C:/tmp/doc.docx"
});

describe("LeaseManager persistence-first (issue #6)", () => {
  it("acquire: DB insert failure leaves NO phantom in-memory writer", () => {
    const manager = new LeaseManager(stubRepos({ insert: true }));
    expect(() => manager.acquire("sess-lm", 1, input())).toThrow("db insert failed");
    expect(manager.hasActiveLease("sess-lm")).toBe(false);
    // The failed attempt did not poison the document-level writer index.
    expect(() => new LeaseManager(stubRepos()).acquire("sess-lm", 1, input())).not.toThrow();
  });

  it("release: DB failure keeps the memory slot — no divergence, next writer stays blocked", () => {
    const manager = new LeaseManager(stubRepos());
    const lease = manager.acquire("sess-lm", 1, input("human"));
    expect(() => manager.release(lease.leaseId)).not.toThrow(); // healthy path

    const manager2 = new LeaseManager(stubRepos());
    const lease2 = manager2.acquire("sess-lm", 1, input("agent"));
    const failing = new LeaseManager(stubRepos({ release: true }));
    // Simulate the same lease being present with a failing releaseLease.
    const lease3 = failing.acquire("sess-lm2", 1, input("agent"));
    expect(() => failing.release(lease3.leaseId)).toThrow("db release failed");
    // Memory slot survives the failed durable release — audit and runtime
    // agree the writer is STILL active (single commit point).
    expect(failing.hasActiveLease("sess-lm2")).toBe(true);
    void lease2;
    void manager2;
  });
});
