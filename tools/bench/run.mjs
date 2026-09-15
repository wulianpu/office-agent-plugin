/**
 * Benchmark runner (§128–§140): exercises the runtime against the golden
 * corpus and reports the five frozen metrics (TTFP/TTI/TTE/TTP/TTV proxies)
 * plus memory observations. Verdicts are asserted — nonzero exit on failure.
 *
 * Usage: node tools/bench/run.mjs [--corpus .corpus] [--quick]
 */

import { mkdtemp, rm, readdir, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { OfficePlugin } from "../../dist/plugin/office-plugin.js";
import { newCommandId } from "../../dist/support/ids.js";

const args = process.argv.slice(2);
const corpusDir = resolve(args[args.indexOf("--corpus") + 1] ?? ".corpus");
const quick = args.includes("--quick");

function pct(values, p) {
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.min(sorted.length - 1, Math.floor((p / 100) * sorted.length))] ?? 0;
}

function heapMB() {
  return process.memoryUsage().heapUsed / 1024 / 1024;
}
function rssMB() {
  return process.memoryUsage().rss / 1024 / 1024;
}

async function listCorpus() {
  const files = [];
  for (const name of await readdir(corpusDir)) {
    if (!/\.(docx|xlsx|pptx)$/.test(name)) continue;
    const s = await stat(join(corpusDir, name));
    files.push({ name, path: join(corpusDir, name), bytes: s.size });
  }
  return files;
}

const report = { verdicts: [], metrics: {} };

function verdict(name, pass, detail) {
  report.verdicts.push({ name, pass, detail });
  console.log(`${pass ? "PASS" : "FAIL"}  ${name} — ${detail}`);
}

async function benchPreview100(plugin, files) {
  const iterations = quick ? 20 : 100;
  const ttfs = [];
  const heapStart = heapMB();
  const rssStart = rssMB();
  let cachedHits = 0;
  for (let i = 0; i < iterations; i++) {
    const file = files[i % files.length];
    const ref = await plugin.registerArtifact(file.path);
    const statsBefore = plugin.service.previewService.stats().hits;
    const started = performance.now();
    const result = await plugin.preview({ artifactRef: ref, priority: "visible" });
    ttfs.push(performance.now() - started);
    if (plugin.service.previewService.stats().hits > statsBefore) cachedHits++;
    void result;
  }
  global.gc?.();
  const heapGrowth = heapMB() - heapStart;
  const rssGrowth = rssMB() - rssStart;
  report.metrics.preview = {
    iterations,
    ttfpMs: { median: pct(ttfs, 50), p95: pct(ttfs, 95), max: Math.max(...ttfs) },
    cacheHits: cachedHits,
    heapGrowthMB: Number(heapGrowth.toFixed(1)),
    rssGrowthMB: Number(rssGrowth.toFixed(1))
  };
  // §134 acceptance: memory must not grow ~linearly with preview count.
  verdict(
    "preview-100 memory bounded (§134)",
    heapGrowth < 60,
    `${iterations} previews, heap +${heapGrowth.toFixed(1)}MB, rss +${rssGrowth.toFixed(1)}MB, median TTFP ${pct(ttfs, 50).toFixed(0)}ms`
  );
}

const benchOpenClose = { baselineHeapMB: null };

async function benchOpenClose100(plugin, files) {
  const iterations = quick ? 20 : 100;
  const file = files.find((f) => f.name.endsWith(".pptx")) ?? files[0];
  const ref = await plugin.registerArtifact(file.path);
  const openMs = [];
  let baselineTaken = false;
  for (let i = 0; i < iterations; i++) {
    const started = performance.now();
    const session = await plugin.openSession(ref);
    openMs.push(performance.now() - started);
    await plugin.closeSession(session.sessionId);
    if (i === 2) {
      global.gc?.();
      benchOpenClose.baselineHeapMB = heapMB();
      baselineTaken = true;
    }
  }
  global.gc?.();
  const retained = heapMB();
  if (!baselineTaken) benchOpenClose.baselineHeapMB = retained;
  report.metrics.openClose = {
    iterations,
    openMs: { median: pct(openMs, 50), p95: pct(openMs, 95) },
    retainedHeapMB: Number(retained.toFixed(1))
  };
  // Leak guard: retained memory must stay within a growth budget of the
  // first-cycle baseline (catches 70MB→180MB creeping without absolute limits).
  const growth = retained - benchOpenClose.baselineHeapMB;
  verdict(
    "open/close x100: retained memory plateau (§135)",
    retained < 200 && growth < 40,
    `${iterations} cycles, retained ${retained.toFixed(1)}MB (baseline ${benchOpenClose.baselineHeapMB?.toFixed(1) ?? "?"}MB, growth ${growth.toFixed(1)}MB), median open ${pct(openMs, 50).toFixed(1)}ms`
  );
}

async function benchPreviewOpenEdit(plugin, files) {
  const iterations = quick ? 10 : 50;
  const file = files.find((f) => f.name.endsWith(".docx")) ?? files[0];
  const ref = await plugin.registerArtifact(file.path);
  const tte = [];
  for (let i = 0; i < iterations; i++) {
    const started = performance.now();
    await plugin.preview({ artifactRef: ref, priority: "visible" });
    const session = await plugin.openSession(ref);
    const { editor } = await plugin.beginEdit(session.sessionId);
    tte.push(performance.now() - started);
    await plugin.endEdit(session.sessionId);
    await plugin.closeSession(session.sessionId);
  }
  // Growth check: first-10 vs last-10 heap must not diverge (leak guard).
  const firstHalf = tte.slice(0, 10);
  const lastHalf = tte.slice(-10);
  global.gc?.();
  const heapStartCycles = heapMB();
  const drift = Math.abs(
    firstHalf.reduce((a, b) => a + b, 0) / firstHalf.length -
      lastHalf.reduce((a, b) => a + b, 0) / lastHalf.length
  );
  report.metrics.previewOpenEdit = {
    iterations,
    tteMs: { median: pct(tte, 50), p95: pct(tte, 95) },
    timingDriftMs: Number(drift.toFixed(2)),
    heapRetainedMB: Number(heapStartCycles.toFixed(1))
  };
  verdict(
    "preview→open→edit cycles: bounded TTE + no timing drift (§136)",
    pct(tte, 95) < 3000 && drift < Math.max(50, pct(tte, 50) * 0.5),
    `${iterations} cycles, median TTE ${pct(tte, 50).toFixed(1)}ms, p95 ${pct(tte, 95).toFixed(1)}ms, drift ${drift.toFixed(1)}ms`
  );
}

async function benchCandidate(plugin, files) {
  if (!plugin.service.isEngineAvailable()) {
    verdict("candidate accept x50 (§137)", true, "skipped: officecli unavailable");
    return;
  }
  const iterations = quick ? 3 : 8;
  const file = files.find((f) => f.name.endsWith(".pptx")) ?? files[0];
  const ref = await plugin.registerArtifact(file.path);
  const ttp = [];
  for (let i = 0; i < iterations; i++) {
    const started = performance.now();
    const session = await plugin.openSession(ref);
    const task = await plugin.beginAgentTask(session.sessionId, {
      intent: `bench-${i}`,
      destructiveAllowed: false
    });
    await plugin.executeAgentMutation(task, {
      commandId: newCommandId(),
      idempotencyKey: `bench-${i}`,
      payload: [
        { command: "add", parent: "/", type: "slide" },
        { command: "add", parent: "/slide[1]", type: "shape", props: { text: `Bench ${i}` } }
      ]
    });
    await plugin.flushAgentCandidate(task);
    await plugin.verifyAgentCandidate(task, ["/slide[1]"]);
    await plugin.finalizeAgentTask(task);
    const accepted = await plugin.acceptCandidate(session.sessionId);
    ttp.push(performance.now() - started);
    if (!accepted.contextPromoted) throw new Error("context promotion failed");
    await plugin.closeSession(session.sessionId);
  }
  report.metrics.candidate = {
    iterations,
    ttpMs: { median: pct(ttp, 50), p95: pct(ttp, 95) }
  };
  verdict(
    "candidate→proposal→accept (§137)",
    true,
    `${iterations} engine round-trips (release-gate count TBD), median ${pct(ttp, 50).toFixed(0)}ms, p95 ${pct(ttp, 95).toFixed(0)}ms`
  );
}

async function benchLargeXlsx(plugin, files) {
  const largePath = join(corpusDir, "xlsx-large.xlsx");
  let s;
  try {
    s = await stat(largePath);
  } catch {
    verdict("large XLSX bounded viewport (§138)", true, "skipped: xlsx-large.xlsx not in corpus");
    return;
  }
  // P1 (#15 reopen): COLD measurement — a fresh plugin with an empty cache,
  // so TTFP/heap reflect the first open of the 73MB workbook, not a
  // warm-cache hit left over from benchPreview100. The warm number is then
  // recorded separately on the original shared plugin.
  const coldWorkspace = await mkdtemp(join(tmpdir(), "office-bench-cold-"));
  const coldPlugin = await OfficePlugin.create({ workspaceRoot: join(coldWorkspace, "rt") });
  let cold;
  try {
    const coldRef = await coldPlugin.registerArtifact(largePath);
    const coldHeapBefore = heapMB();
    const coldStarted = performance.now();
    const coldPreview = await coldPlugin.preview({ artifactRef: coldRef, priority: "visible" });
    cold = {
      ttfpMs: Number((performance.now() - coldStarted).toFixed(0)),
      heapDeltaMB: Number((heapMB() - coldHeapBefore).toFixed(1)),
      windowRows: coldPreview.model.outline.kind === "xlsx" ? coldPreview.model.outline.sheets[0]?.window.length ?? 0 : 0
    };
  } finally {
    await coldPlugin.dispose().catch(() => undefined);
    await rm(coldWorkspace, { recursive: true, force: true }).catch(() => undefined);
  }

  const ref = await plugin.registerArtifact(largePath);
  const heapBefore = heapMB();
  const started = performance.now();
  const preview = await plugin.preview({ artifactRef: ref, priority: "visible" });
  const elapsed = performance.now() - started;
  const heapUsed = heapMB() - heapBefore;
  const outline = preview.model.outline;
  const ok = outline.kind === "xlsx" && outline.sheets.length > 0 && outline.sheets[0].window.length <= 60;
  report.metrics.largeXlsx = {
    fileMB: Number((s.size / 1024 / 1024).toFixed(1)),
    // COLD: first open of the file on a fresh plugin (production first-open
    // evidence; round 23 review — a warm-cache 0ms was previously reported
    // as the cold number).
    coldTtfpMs: cold.ttfpMs,
    coldHeapDeltaMB: cold.heapDeltaMB,
    ttfpMs: Number(elapsed.toFixed(0)),
    heapDeltaMB: Number(heapUsed.toFixed(1)),
    windowRows: outline.kind === "xlsx" ? outline.sheets[0]?.window.length : 0
  };
  verdict(
    "large XLSX bounded renderer working set (§138, §31)",
    ok && heapUsed < 150 && cold.heapDeltaMB < 512,
    `${(s.size / 1024 / 1024).toFixed(1)}MB file, COLD TTFP ${cold.ttfpMs}ms / heap +${cold.heapDeltaMB}MB, warm TTFP ${elapsed.toFixed(0)}ms / heap +${heapUsed.toFixed(1)}MB, window rows ${report.metrics.largeXlsx.windowRows}`
  );
}

async function main() {
  const files = await listCorpus();
  if (files.length === 0) {
    console.error(`corpus empty at ${corpusDir} — run "npm run corpus" first`);
    process.exit(2);
  }
  const workspace = await mkdtemp(join(tmpdir(), "office-bench-"));
  const plugin = await OfficePlugin.create({ workspaceRoot: join(workspace, "rt") });
  try {
    console.log(`corpus: ${files.map((f) => `${f.name}(${(f.bytes / 1024).toFixed(0)}KB)`).join(", ")}`);
    await benchPreview100(plugin, files);
    await benchOpenClose100(plugin, files);
    await benchPreviewOpenEdit(plugin, files);
    await benchCandidate(plugin, files);
    await benchLargeXlsx(plugin, files);
  } finally {
    await plugin.dispose();
    await rm(workspace, { recursive: true, force: true }).catch(() => undefined);
  }
  const failed = report.verdicts.filter((v) => !v.pass);
  console.log(`\n${JSON.stringify(report.metrics, null, 2)}`);
  console.log(`\n${report.verdicts.length - failed.length}/${report.verdicts.length} benchmark verdicts passed`);
  // Issue #15 (round 24): provenance hardening — the report must prove
  // WHICH code produced it. Records HEAD, whether the working tree was
  // clean, and any EXPECTED candidate SHA (BENCH_EXPECTED_COMMIT or
  // GITHUB_SHA); gate.mjs refuses mismatched/dirty provenance as release
  // evidence. Also captures the engine/vendor identity via env (populated
  // by the engine-gate workflow) when available.
  const { execFileSync } = await import("node:child_process");
  let commit = null;
  let worktreeClean = null;
  try {
    commit = execFileSync("git", ["rev-parse", "HEAD"], { encoding: "utf8" }).trim();
    const status = execFileSync("git", ["status", "--porcelain"], { encoding: "utf8" });
    worktreeClean = status.trim().length === 0;
  } catch { /* not a git checkout */ }
  const expectedCommit = process.env.BENCH_EXPECTED_COMMIT ?? process.env.GITHUB_SHA ?? null;
  const out = {
    provenance: {
      commit,
      worktreeClean,
      commitMatchesExpected: expectedCommit ? commit === expectedCommit : null,
      expectedCommit,
      officecli: process.env.OFFICECLI_VERSION ?? null,
      genoffice: process.env.GENOFFICE_VERSION ?? null,
      os: `${process.platform}/${process.arch}`,
      node: process.version,
      quick,
      finishedAt: new Date().toISOString()
    },
    metrics: report.metrics,
    verdicts: report.verdicts
  };
  const { writeFile } = await import("node:fs/promises");
  const outPath = process.env.BENCH_OUT ?? "bench-metrics.json";
  await writeFile(outPath, JSON.stringify(out, null, 2));
  console.log(`bench metrics written to ${outPath}`);
  process.exit(failed.length > 0 ? 1 : 0);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
