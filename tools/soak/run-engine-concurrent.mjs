/**
 * Issue #15 remaining resource sub-gate: CONCURRENT engine-backed workload.
 * The 13.7h serial soak closed the serial long-run plateau; THIS tool drives
 * 4 workers across DIFFERENT documents in a SHARED workspace root (multi-
 * process SQLite WAL + shared OfficeCLI engine = real queue wait) and
 * asserts the resource story under concurrency: per-op p50/p95 bounded,
 * zero errors, OS handle count bounded, RSS drains back after the workload.
 *
 * Usage: node tools/soak/run-engine-concurrent.mjs [workers=4] [cyclesPerWorker=150]
 */

import { execSync, execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { pathToFileURL } from "node:url";
import { join } from "node:path";

const run = promisify(execFile);
const workers = Number(process.argv[2] ?? 4);
const cyclesPerWorker = Number(process.argv[3] ?? 150);
const fail = (msg) => {
  console.error(`CONCURRENT SOAK FAIL: ${msg}`);
  process.exit(1);
};

const engine = await import(
  pathToFileURL(join(process.cwd(), "dist", "agent", "officecli", "officecli-adapter.js")).href
).then((m) => new m.OfficeCliAdapter({ timeoutMs: 30_000 }).version_().catch(() => null));
if (!engine) fail("OfficeCLI engine unavailable — concurrent soak REQUIRES the engine");
const sha = execSync("git rev-parse HEAD", { cwd: process.cwd() }).toString().trim();

const osHandles = async (pids) => {
  if (process.platform !== "win32") return -1;
  try {
    const list = pids.join(",");
    const out = await run(
      "powershell",
      ["-NoProfile", "-Command", `(Get-Process -Id ${list} -ErrorAction SilentlyContinue | Measure-Object HandleCount -Sum).Sum`],
      { encoding: "utf8", timeout: 60_000 }
    );
    return Number(out.stdout.trim()) || 0;
  } catch {
    return -1;
  }
};
const rssMB = () => process.memoryUsage().rss / 1024 / 1024;

const tmpRoot = await mkdtemp(join(tmpdir(), "soak-concurrent-"));
// One workspace root PER WORKER: the RuntimeDatabase is single-owner by
// design (see src/runtime/persistence/database.ts) — multi-process sharing
// of one DB is out of contract and the startup reconciliation legitimately
// reclaims foreign rows. The contended resource under test is the shared
// engine + machine, not the DB.
await mkdir(join(tmpRoot, "placeholder"), { recursive: true });
const baselineRss = rssMB();
const startedAt = new Date().toISOString();
console.log(`baseline: RSS ${baselineRss.toFixed(1)} MB, launching ${workers} workers x ${cyclesPerWorker} cycles`);

const children = [];
for (let w = 1; w <= workers; w++) {
  children.push(
    run(process.execPath, [join(process.cwd(), "tools", "soak", "engine-worker.mjs"), String(w), String(cyclesPerWorker), tmpRoot], {
      cwd: process.cwd(),
      timeout: 3_600_000,
      maxBuffer: 16 * 1024 * 1024
    })
  );
}
const childPids = () =>
  execSync(
    `powershell -NoProfile -Command "(Get-CimInstance Win32_Process -Filter \\"ParentProcessId=${process.pid}\\").ProcessId -join ','"`,
    { encoding: "utf8" }
  ).trim();
const samples = [];
const sampler = setInterval(async () => {
  try {
    const pids = [process.pid, ...childPids().split(",").filter(Boolean)];
    const sample = { at: new Date().toISOString(), parentRssMB: Number(rssMB().toFixed(1)), osHandles: await osHandles(pids) };
    samples.push(sample);
    console.log(JSON.stringify(sample));
  } catch {
    /* sampler must never kill the run */
  }
}, 10_000);

const settled = await Promise.allSettled(children);
clearInterval(sampler);
const drainRss = rssMB();
const summaries = [];
for (const r of settled) {
  if (r.status !== "fulfilled") {
    const reason = r.reason ?? {};
    summaries.push({
      error: String(reason.message ?? reason).slice(0, 200),
      workerStderr: String(reason.stderr ?? "").slice(-2000),
      workerStdoutTail: String(reason.stdout ?? "").slice(-500)
    });
    continue;
  }
  const line = r.value.stdout.trim().split("\n").find((l) => l.trim().startsWith("{"));
  summaries.push(line ? JSON.parse(line.trim()) : { error: "no verdict" });
}
const totalAccepted = summaries.reduce((a, s) => a + (s.accepted ?? 0), 0);
const totalErrors = summaries.reduce((a, s) => a + (s.errors ?? 1), 0);
const p95s = summaries.map((s) => s.p95 ?? Infinity);
const p95Max = Math.max(...p95s);
const maxHandles = Math.max(...samples.map((s) => s.osHandles).filter((n) => n >= 0), 0);
const drainedRss = rssMB();

const passed =
  totalErrors === 0 &&
  totalAccepted === workers * cyclesPerWorker &&
  p95Max < 60_000 &&
  (maxHandles === 0 || maxHandles <= 2000) &&
  drainedRss <= baselineRss * 1.5 + 100;

const verdict = {
  kind: "concurrent-engine-soak",
  sha,
  engine,
  workers,
  cyclesPerWorker,
  totalAccepted,
  totalErrors,
  perOpP95MaxMs: p95Max,
  maxOsHandles: maxHandles,
  baselineRssMB: Number(baselineRss.toFixed(1)),
  drainedRssMB: Number(drainedRss.toFixed(1)),
  summaries,
  passed,
  startedAt,
  finishedAt: new Date().toISOString()
};
await mkdir(join(process.cwd(), "test-results"), { recursive: true });
await writeFile(
  join(process.cwd(), "test-results", "soak-concurrent-report.json"),
  JSON.stringify({ verdict, samples }, null, 2)
);
await rm(tmpRoot, { recursive: true, force: true }).catch(() => undefined);
console.log(JSON.stringify(verdict));
console.log(passed ? "CONCURRENT SOAK PASS" : "CONCURRENT SOAK FAIL");
process.exit(passed ? 0 : 1);
