/**
 * Production readiness gate (issue #15 P1): hard-limit smoke wrapper over
 * tools/bench/run.mjs output. The bench already asserts its own verdicts
 * (nonzero exit on failure); this gate adds cross-workload ABSOLUTE limits
 * so a leak / latency / viewport runaway fails CI even when bench verdicts
 * stay green. Deliberately loose: "clearly broken" detector — baseline
 * comparison belongs to the nightly/full bench (issue #15 P1, tier 2).
 *
 * Usage: node tools/bench/gate.mjs <bench-metrics-json>
 * Exits non-zero on any hard-limit violation.
 */

import { readFileSync } from "node:fs";

const jsonPath = process.argv[2];
if (!jsonPath) {
  console.error("usage: node tools/bench/gate.mjs <bench-metrics.json>");
  process.exit(2);
}

// Hard smoke limits — "clearly broken" detectors (issue #15, smoke tier).
const LIMITS = {
  /** ms: p95 preview TTFP must stay bounded (viewport runaway / leak). */
  previewTtfpP95Ms: 30_000,
  /** MB: preview-100 heap growth must stay bounded (linear-growth leak). */
  previewHeapGrowthMB: 200,
  /** ms: p95 open→first-edit activation. */
  previewOpenEditTteP95Ms: 60_000,
  /** ms: p95 candidate→proposal→accept engine round-trip. */
  candidateTtpP95Ms: 120_000,
  /** MB: large-XLSX bounded-viewport heap delta (runaway window). */
  largeXlsxHeapDeltaMB: 512,
  /** Round 23: COLD first-open bounds for the 73MB workbook. */
  largeXlsxColdTtfpMs: 60_000,
  largeXlsxColdHeapMB: 1024
};

function fail(metric, detail) {
  console.error(`GATE FAIL: ${metric} — ${detail}`);
  process.exit(1);
}

/**
 * P1 (#15 reopen): mandatory-metric fail-closed. Every workload metric in
 * this list MUST be present and finite in the report — an absent/NaN/
 * Infinity mandatory metric fails the gate instead of silently skipping
 * the check (a missing candidate/largeXlsx section previously produced a
 * green "GATE PASS" with those workloads effectively unexecuted).
 */
const MANDATORY = [
  ["preview", "ttfpMs", "p95", LIMITS.previewTtfpP95Ms],
  ["preview", "heapGrowthMB", null, LIMITS.previewHeapGrowthMB],
  ["previewOpenEdit", "tteMs", "p95", LIMITS.previewOpenEditTteP95Ms],
  ["candidate", "ttpMs", "p95", LIMITS.candidateTtpP95Ms],
  ["largeXlsx", "coldTtfpMs", null, LIMITS.largeXlsxColdTtfpMs],
  ["largeXlsx", "coldHeapDeltaMB", null, LIMITS.largeXlsxColdHeapMB],
  ["largeXlsx", "heapDeltaMB", null, LIMITS.largeXlsxHeapDeltaMB]
];

const report = JSON.parse(readFileSync(jsonPath, "utf8"));

// P1 (#15, round 25): provenance enforcement.
//  - expected-SHA mode (BENCH_EXPECTED_COMMIT / GITHUB_SHA present): the
//    report commit must MATCH and the producing worktree must have been
//    CLEAN — a dirty tree means the executed bytes were not the candidate
//    bytes even when HEAD === expected.
//  - expected-SHA absent: local ad-hoc run — provenance is still recorded
//    but the report is NOT release evidence (RELEASE.md §1).
const prov = report.provenance ?? {};
const expected = process.env.BENCH_EXPECTED_COMMIT ?? process.env.GITHUB_SHA;
if (expected) {
  if (prov.commit !== expected) {
    console.error(
      `GATE FAIL: provenance commit mismatch — report from ${prov.commit ?? "?"}, expected ${expected}`
    );
    process.exit(1);
  }
  if (prov.worktreeClean !== true) {
    console.error(
      "GATE FAIL: provenance worktreeClean is not true — the benchmark executed against uncommitted changes and cannot serve as release evidence"
    );
    process.exit(1);
  }
  if (prov.commitMatchesExpected !== true) {
    console.error("GATE FAIL: provenance commitMatchesExpected is not true");
    process.exit(1);
  }
}

const m = report.metrics ?? report;
if (!m || typeof m !== "object" || Object.keys(m).length === 0) {
  fail("bench-report", "no metrics in report — check bench wiring");
}

const violations = [];

function readMetric(section, leaf, nestedKey) {
  const obj = m[section];
  if (!obj || typeof obj !== "object") return undefined;
  if (nestedKey) {
    const inner = obj[leaf];
    return inner ? inner[nestedKey] : undefined;
  }
  return obj[leaf];
}

for (const [section, leaf, nestedKey, limit] of MANDATORY) {
  const path = nestedKey ? `${section}.${leaf}.${nestedKey}` : `${section}.${leaf}`;
  const v = readMetric(section, leaf, nestedKey);
  if (typeof v !== "number" || !Number.isFinite(v)) {
    violations.push([path, `mandatory metric missing or non-finite (${String(v)}) — workload must execute`]);
  }
}

if (m.preview) {
  const v = m.preview.ttfpMs?.p95;
  if (typeof v === "number" && Number.isFinite(v) && v > LIMITS.previewTtfpP95Ms) {
    violations.push(["preview.ttfpMs.p95", `${v}ms > ${LIMITS.previewTtfpP95Ms}ms`]);
  }
}
if (m.preview?.heapGrowthMB !== undefined) {
  const v = m.preview.heapGrowthMB;
  if (v > LIMITS.previewHeapGrowthMB) violations.push(["preview.heapGrowthMB", `${v}MB > ${LIMITS.previewHeapGrowthMB}MB`]);
}
if (m.previewOpenEdit?.tteMs?.p95 !== undefined) {
  const v = m.previewOpenEdit.tteMs.p95;
  if (v > LIMITS.previewOpenEditTteP95Ms) violations.push(["previewOpenEdit.tteMs.p95", `${v}ms > ${LIMITS.previewOpenEditTteP95Ms}ms`]);
}
if (m.candidate?.ttpMs?.p95 !== undefined) {
  const v = m.candidate.ttpMs.p95;
  if (v > LIMITS.candidateTtpP95Ms) violations.push(["candidate.ttpMs.p95", `${v}ms > ${LIMITS.candidateTtpP95Ms}ms`]);
}
if (m.largeXlsx?.heapDeltaMB !== undefined) {
  const v = m.largeXlsx.heapDeltaMB;
  if (v > LIMITS.largeXlsxHeapDeltaMB) violations.push(["largeXlsx.heapDeltaMB", `${v}MB > ${LIMITS.largeXlsxHeapDeltaMB}MB`]);
}
if (m.largeXlsx?.coldTtfpMs !== undefined) {
  const v = m.largeXlsx.coldTtfpMs;
  if (v > LIMITS.largeXlsxColdTtfpMs) violations.push(["largeXlsx.coldTtfpMs", `${v}ms > ${LIMITS.largeXlsxColdTtfpMs}ms`]);
}
if (m.largeXlsx?.coldHeapDeltaMB !== undefined) {
  const v = m.largeXlsx.coldHeapDeltaMB;
  if (v > LIMITS.largeXlsxColdHeapMB) violations.push(["largeXlsx.coldHeapDeltaMB", `${v}MB > ${LIMITS.largeXlsxColdHeapMB}MB`]);
}

if (violations.length > 0) {
  for (const [metric, detail] of violations) console.error(`GATE FAIL: ${metric} — ${detail}`);
  process.exit(1);
}

console.log("GATE PASS: all mandatory metrics present, finite, and within hard limits");
