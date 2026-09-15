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
  largeXlsxHeapDeltaMB: 512
};

function fail(metric, detail) {
  console.error(`GATE FAIL: ${metric} — ${detail}`);
  process.exit(1);
}

const report = JSON.parse(readFileSync(jsonPath, "utf8"));
const m = report.metrics ?? report;
if (!m || typeof m !== "object" || Object.keys(m).length === 0) {
  fail("bench-report", "no metrics in report — check bench wiring");
}

const violations = [];

if (m.preview?.ttfpMs?.p95 !== undefined) {
  const v = m.preview.ttfpMs.p95;
  if (v > LIMITS.previewTtfpP95Ms) violations.push(["preview.ttfpMs.p95", `${v}ms > ${LIMITS.previewTtfpP95Ms}ms`]);
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

if (violations.length > 0) {
  for (const [metric, detail] of violations) console.error(`GATE FAIL: ${metric} — ${detail}`);
  process.exit(1);
}

console.log("GATE PASS: all hard smoke limits satisfied");
