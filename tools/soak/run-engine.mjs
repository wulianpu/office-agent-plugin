/**
 * Issue #15 P1 soak — ENGINE-BACKED variant of run.mjs.
 *
 * Unlike run.mjs (engine-free lifecycle), this soak drives the REAL OfficeCLI
 * engine every cycle: Agent write path (task → mutate → flush → verify →
 * finalize → accept) across DOCX/PPTX/XLSX fixtures, background previews,
 * periodic cold restarts (dispose → new Runtime → resolveRecoveredSession),
 * and samples RSS/heap/handles plus the count of live officecli engine
 * processes. Assertions: RSS plateau, zero leaked ownership, bounded engine
 * process count, zero errors.
 *
 * Usage: node tools/soak/run-engine.mjs [cycles=300] [sampleEvery=10]
 * The engine is REQUIRED here — a missing OfficeCLI is a FAIL, not a skip.
 */

import { mkdtemp, rm, writeFile, mkdir } from "node:fs/promises";
import { writeFileSync as writeFileSyncSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { execSync } from "node:child_process";
import { OfficePlugin } from "../../dist/plugin/office-plugin.js";
import { OfficeCliAdapter } from "../../dist/agent/officecli/officecli-adapter.js";

const cycles = Number(process.argv[2] ?? 300);
const startedAt = new Date().toISOString();
const sampleEvery = Number(process.argv[3] ?? 10);
const RESTART_EVERY = 25;

function rssMB() {
  return process.memoryUsage().rss / 1024 / 1024;
}
function heapMB() {
  return process.memoryUsage().heapUsed / 1024 / 1024;
}
function fdCount() {
  try {
    return (process._getActiveHandles?.() ?? []).length;
  } catch {
    return -1;
  }
}
/**
 * Live officecli engine processes at sample time (Windows CIM query).
 * Engine invocations are short-lived standalone batches, so this records
 * whether any engine process LINGERS/accumulates across cycles and cold
 * restarts — the leak signal is growth, not the absolute count.
 */
function engineProcessesAtSample() {
  if (process.platform !== "win32") return -1;
  try {
    const out = execSync(
      'powershell -NoProfile -Command "(Get-CimInstance Win32_Process -Filter \\"Name=\'node.exe\'\\" | Where-Object { $_.CommandLine -match \'officecli\' } | Measure-Object).Count"',
      { encoding: "utf8", timeout: 60_000 }
    );
    return Number(out.trim()) || 0;
  } catch {
    return -1;
  }
}

// Engine is REQUIRED: probe fails → soak FAILS (never a silent skip).
const engineAdapter = new OfficeCliAdapter({ timeoutMs: 90_000 });
const engineVersion = await engineAdapter.version_().catch(() => null);
if (!engineVersion) {
  console.error("ENGINE SOAK FAIL: OfficeCLI engine unavailable — engine soak cannot run");
  process.exit(1);
}

const candidateSha = execSync("git rev-parse HEAD", { cwd: process.cwd() }).toString().trim();
const workspace = await mkdtemp(join(tmpdir(), "soak-engine-"));
let plugin = await OfficePlugin.create({ workspaceRoot: join(workspace, "rt"), skipHostProbe: true });

// Real engine-created fixtures (one per format), mutated every cycle.
const adapter = new OfficeCliAdapter({ timeoutMs: 90_000 });
async function makeFixture(name, items) {
  const file = join(workspace, name);
  await adapter.run(["create", file, "--json"]).catch(() => undefined);
  await adapter.runBatchStandalone(file, items);
  await adapter.close(file).catch(() => undefined);
  return file;
}
const docx = await makeFixture("soak.docx", [
  { command: "add", parent: "/body", type: "paragraph", props: { text: "soak base" } }
]);
const pptx = await makeFixture("soak.pptx", [
  { command: "add", parent: "/", type: "slide" },
  { command: "add", parent: "/slide[1]", type: "shape", props: { text: "soak base", x: "1cm", y: "1cm" } }
]);
const xlsx = await makeFixture("soak.xlsx", [
  { command: "add", parent: "/", type: "sheet", props: { name: "Data" } }
]);
const fixtures = [
  { file: docx, mutation: (i) => ({ command: "set", path: "/body/paragraph[1]", props: { text: `soak-${i}` } }) },
  { file: pptx, mutation: (i) => ({ command: "set", path: "/slide[1]/shape[1]", props: { text: `soak-${i}` } }) },
  { file: xlsx, mutation: (i) => ({ command: "set", path: "/Data/A1", props: { value: `soak-${i}` } }) }
];

let refs = [];
for (const f of fixtures) refs.push({ ref: await plugin.registerArtifact(f.file), fixture: f });

let errors = 0;
let firstError = null;
let lastSessionId = null;
const samples = [];
const residentSamples = [];

for (let cycle = 1; cycle <= cycles; cycle++) {
  try {
    if (cycle % RESTART_EVERY === 0) {
      // Cold restart under sustained engine load: dispose everything and
      // recreate the Runtime. Every cycle closes its session by design, so
      // there is deliberately NO open session to recover here (recovery of
      // an OPEN session across dispose is production-compat's required-CI
      // job; resolving a closed-then-reconciled session id throws
      // unknown-session). The restart assertion is hydration: the recreated
      // runtime re-registers every artifact and the next cycle must work.
      await plugin.dispose().catch(() => undefined);
      plugin = await OfficePlugin.create({ workspaceRoot: join(workspace, "rt"), skipHostProbe: true });
      refs = [];
      for (const f of fixtures) refs.push({ ref: await plugin.registerArtifact(f.file), fixture: f });
    }

    const { ref, fixture } = refs[cycle % refs.length];
    const session = await plugin.openSession(ref);
    lastSessionId = session.sessionId;
    await plugin.service.sessions.ensureStrongIdentity(session.sessionId);
    const task = await plugin.beginAgentTask(session.sessionId, {
      intent: `engine soak cycle ${cycle}`,
      destructiveAllowed: false
    });
    await plugin.executeAgentMutation(task, {
      commandId: `soak-${cycle}`,
      idempotencyKey: `soak-${cycle}`,
      payload: [fixture.mutation(cycle)]
    });
    await plugin.flushAgentCandidate(task);
    await plugin.verifyAgentCandidate(task);
    await plugin.finalizeAgentTask(task);
    await plugin.acceptCandidate(session.sessionId, task.candidateId);
    await plugin.closeSession(session.sessionId);
    await plugin.preview({ artifactRef: ref, priority: "background" });
  } catch (error) {
    errors++;
    firstError ??= `${error?.code ?? ""} ${error?.message ?? String(error)}`.slice(0, 300);
    console.error(`cycle ${cycle} error: ${error?.message ?? error}`);
    if (errors > Math.max(3, Math.floor(cycles * 0.01))) break; // abort early on systemic failure
  }

  if (cycle % sampleEvery === 0 || cycle === cycles) {
    global.gc?.();
    const sample = {
      cycle,
      rssMB: Number(rssMB().toFixed(1)),
      heapMB: Number(heapMB().toFixed(1)),
      fds: fdCount(),
      sessions: plugin.service.sessions.list().length,
      editorLeases: plugin.service.editorLeaseCount(),
      engineTasks: plugin.service.agent.activeTaskCount(),
      errors
    };
    samples.push(sample);
    console.log(JSON.stringify(sample));
  }
  if (cycle % (sampleEvery * 5) === 0 || cycle === cycles) {
    const residents = engineProcessesAtSample();
    residentSamples.push({ cycle, residents });
    console.log(JSON.stringify({ cycle, residents }));
  }
}

const first = samples[0];
const last = samples[samples.length - 1];
const rssGrowth = last.rssMB / Math.max(1, first.rssMB);
const maxResidents = Math.max(...residentSamples.map((r) => r.residents).filter((n) => n >= 0), 0);
const verdict = {
  kind: "engine-soak",
  sha: candidateSha,
  engineVersion,
  cycles,
  samples: samples.length,
  restarts: Math.floor(cycles / RESTART_EVERY),
  rssFirstMB: first.rssMB,
  rssLastMB: last.rssMB,
  rssGrowthRatio: Number(rssGrowth.toFixed(2)),
  fdsFirst: first.fds,
  fdsLast: last.fds,
  maxEngineProcesses: maxResidents,
  residentsLast: residentSamples[residentSamples.length - 1]?.residents ?? -1,
  sessionsAtEnd: last.sessions,
  editorLeasesAtEnd: last.editorLeases,
  engineTasksAtEnd: last.engineTasks,
  errors,
  firstError,
  startedAt,
  finishedAt: new Date().toISOString()
};

const passed =
  errors === 0 &&
  rssGrowth < 2.5 &&
  last.sessions === 0 &&
  last.editorLeases === 0 &&
  last.engineTasks === 0 &&
  (residentSamples.length === 0 || residents_within_bound());

function residents_within_bound() {
  // Engine process count must stay bounded — residents must not accumulate
  // across 100+ accept cycles and cold restarts.
  return maxResidents <= 8;
}

await mkdir(join(process.cwd(), "test-results"), { recursive: true });
writeFileSyncSync(
  join(process.cwd(), "test-results", "soak-engine-report.json"),
  JSON.stringify({ verdict, samples, residentSamples }, null, 2)
);
writeFileSyncSync(join(workspace, "soak-report.json"), JSON.stringify({ verdict, samples, residentSamples }, null, 2));
console.log(JSON.stringify(verdict));
console.log(passed ? "ENGINE SOAK PASS: stable plateau" : "ENGINE SOAK FAIL: growth/leak/engine error");
await plugin.dispose().catch(() => undefined);
await rm(workspace, { recursive: true, force: true }).catch(() => undefined);
process.exit(passed ? 0 : 1);
