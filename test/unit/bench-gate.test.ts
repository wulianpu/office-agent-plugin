/**
 * Issue #15 (round 25): bench gate provenance enforcement — four-quadrant
 * verification of expected-SHA × worktreeClean, plus mandatory-metric
 * fail-closed regression for the new cold/warm largeXlsx metrics.
 */

import { describe, expect, it } from "vitest";
import { writeFileSync } from "node:fs";
import { execFileSync } from "node:child_process";

const GATE = "tools/bench/gate.mjs";
const EXPECTED = "eb303b1acee1dcb535e2c4bc2ce5f1e237fdce41";

const FULL_METRICS = {
  preview: { ttfpMs: { p95: 1200 }, heapGrowthMB: 12 },
  previewOpenEdit: { tteMs: { p95: 800 } },
  candidate: { ttpMs: { p95: 9000 } },
  largeXlsx: { coldTtfpMs: 1053, coldHeapDeltaMB: 0.4, heapDeltaMB: 30 }
};

function report(commit: string, opts: { worktreeClean?: boolean; commitMatchesExpected?: boolean } = {}) {
  return {
    provenance: {
      commit,
      worktreeClean: opts.worktreeClean ?? true,
      commitMatchesExpected: opts.commitMatchesExpected ?? true,
      os: "win32/x64",
      node: "v24.20.0",
      finishedAt: new Date().toISOString()
    },
    metrics: FULL_METRICS,
    verdicts: []
  };
}

function runGate(reportObj: unknown, env: Record<string, string>): { code: number; output: string } {
  const path = "bench-gate-probe.json";
  writeFileSync(path, JSON.stringify(reportObj));
  try {
    const out = execFileSync(process.execPath, [GATE, path], {
      env: { ...process.env, ...env },
      encoding: "utf8"
    });
    return { code: 0, output: out };
  } catch (error) {
    const e = error as { status?: number; stdout?: string; stderr?: string };
    return { code: e.status ?? 1, output: `${e.stdout ?? ""}${e.stderr ?? ""}` };
  } finally {
    const { rmSync } = require("node:fs") as typeof import("node:fs");
    rmSync(path, { force: true });
  }
}

describe("bench gate provenance enforcement (issue #15, round 25)", () => {
  it("clean worktree + matching expected SHA → PASS", () => {
    const r = runGate(report(EXPECTED), { BENCH_EXPECTED_COMMIT: EXPECTED });
    expect(r.code).toBe(0);
    expect(r.output).toContain("GATE PASS");
  });

  it("dirty worktree (worktreeClean=false) with matching SHA → FAIL", () => {
    const r = runGate(
      report(EXPECTED, { worktreeClean: false }),
      { BENCH_EXPECTED_COMMIT: EXPECTED }
    );
    expect(r.code).not.toBe(0);
    expect(r.output).toContain("worktreeClean");
  });

  it("HEAD != expected SHA → FAIL (existing behavior preserved)", () => {
    const r = runGate(
      report("deadbeef".repeat(5).slice(0, 40)),
      { BENCH_EXPECTED_COMMIT: EXPECTED }
    );
    expect(r.code).not.toBe(0);
    expect(r.output).toContain("provenance commit mismatch");
  });

  it("commitMatchesExpected=false is refused even with clean worktree", () => {
    const r = runGate(
      report(EXPECTED, { commitMatchesExpected: false }),
      { BENCH_EXPECTED_COMMIT: EXPECTED }
    );
    expect(r.code).not.toBe(0);
  });

  it("no expected SHA provided: local ad-hoc run passes on provenance alone", () => {
    const r = runGate(report("c0ffee".padEnd(40, "0")), {});
    expect(r.code).toBe(0);
  });

  it("mandatory metric missing (candidate section absent) → FAIL", () => {
    const metrics = JSON.parse(JSON.stringify(FULL_METRICS));
    delete (metrics as Record<string, unknown>).candidate;
    const r = runGate(
      { provenance: { commit: EXPECTED, worktreeClean: true, commitMatchesExpected: true }, metrics },
      { BENCH_EXPECTED_COMMIT: EXPECTED }
    );
    expect(r.code).not.toBe(0);
    expect(r.output).toContain("candidate.ttpMs.p95");
  });
});
