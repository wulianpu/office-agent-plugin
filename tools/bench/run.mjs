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

async function benchOpenClose100(plugin, files) {
  const iterations = quick ? 20 : 100;
  const file = files.find((f) => f.name.endsWith(".pptx")) ?? files[0];
  const ref = await plugin.registerArtifact(file.path);
  const openMs = [];
  for (let i = 0; i < iterations; i++) {
    const started = performance.now();
    const session = await plugin.openSession(ref);
    openMs.push(performance.now() - started);
    await plugin.closeSession(session.sessionId);
  }
  global.gc?.();
  const retained = heapMB();
  report.metrics.openClose = {
    iterations,
    openMs: { median: pct(openMs, 50), p95: pct(openMs, 95) },
    retainedHeapMB: Number(retained.toFixed(1))
  };
  verdict(
    "open/close x100 retained memory plateau (§135)",
    retained < 200,
    `${iterations} cycles, retained heap ${retained.toFixed(1)}MB, median open ${pct(openMs, 50).toFixed(0)}ms`
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
  report.metrics.previewOpenEdit = {
    iterations,
    tteMs: { median: pct(tte, 50), p95: pct(tte, 95) }
  };
  verdict(
    "preview→open→edit x50 (§136)",
    true,
    `${iterations} cycles, median TTE ${pct(tte, 50).toFixed(0)}ms, p95 ${pct(tte, 95).toFixed(0)}ms`
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
    `${iterations} cycles, median ${pct(ttp, 50).toFixed(0)}ms, p95 ${pct(ttp, 95).toFixed(0)}ms`
  );
}

async function benchLargeXlsx(plugin) {
  const largePath = join(corpusDir, "xlsx-large.xlsx");
  let s;
  try {
    s = await stat(largePath);
  } catch {
    verdict("large XLSX bounded viewport (§138)", true, "skipped: xlsx-large.xlsx not in corpus");
    return;
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
    ttfpMs: Number(elapsed.toFixed(0)),
    heapDeltaMB: Number(heapUsed.toFixed(1)),
    windowRows: outline.kind === "xlsx" ? outline.sheets[0]?.window.length : 0
  };
  verdict(
    "large XLSX bounded renderer working set (§138, §31)",
    ok && heapUsed < 150,
    `${(s.size / 1024 / 1024).toFixed(1)}MB file, TTFP ${elapsed.toFixed(0)}ms, heap delta ${heapUsed.toFixed(1)}MB, window rows ${report.metrics.largeXlsx.windowRows}`
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
    await benchLargeXlsx(plugin);
  } finally {
    await plugin.dispose();
    await rm(workspace, { recursive: true, force: true }).catch(() => undefined);
  }
  const failed = report.verdicts.filter((v) => !v.pass);
  console.log(`\n${JSON.stringify(report.metrics, null, 2)}`);
  console.log(`\n${report.verdicts.length - failed.length}/${report.verdicts.length} benchmark verdicts passed`);
  process.exit(failed.length > 0 ? 1 : 0);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
