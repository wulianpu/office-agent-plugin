/**
 * One concurrent-soak worker (child process): drives the REAL OfficeCLI
 * engine Agent accept path on its OWN document inside a SHARED workspace
 * root — multi-process SQLite WAL + shared engine = genuine contention and
 * queue wait. Prints one JSON summary line on completion.
 *
 * Usage: node tools/soak/engine-worker.mjs <workerId> <cycles> <workspaceRoot>
 */

import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

const [workerId, cyclesArg, tmpRoot] = process.argv.slice(2);
const workspaceRoot = join(tmpRoot, `ws-w${workerId}`);
const cycles = Number(cyclesArg ?? 100);

const mod = await import(
  pathToFileURL(join(process.cwd(), "dist", "plugin", "office-plugin.js")).href
);
const adapterMod = await import(
  pathToFileURL(join(process.cwd(), "dist", "agent", "officecli", "officecli-adapter.js")).href
);

await (await import("node:fs/promises")).mkdir(workspaceRoot, { recursive: true });
const engine = new adapterMod.OfficeCliAdapter({ timeoutMs: 120_000 });
const docx = join(workspaceRoot, `concurrent-w${workerId}.docx`);
await engine.run(["create", docx, "--json"]).catch(() => undefined);
await engine
  .runBatchStandalone(docx, [{ command: "add", parent: "/body", type: "paragraph", props: { text: `worker ${workerId} base` } }])
  .catch(() => undefined);
await engine.close(docx).catch(() => undefined);

const plugin = await mod.OfficePlugin.create({ workspaceRoot, skipHostProbe: true });
const ref = await plugin.registerArtifact(docx);
let errors = 0;
let firstError = null;
const durations = [];

for (let cycle = 1; cycle <= cycles; cycle++) {
  const t0 = Date.now();
  try {
    const session = await plugin.openSession(ref);
    await plugin.service.sessions.ensureStrongIdentity(session.sessionId);
    const task = await plugin.beginAgentTask(session.sessionId, {
      intent: `concurrent soak w${workerId} cycle ${cycle}`,
      destructiveAllowed: false
    });
    await plugin.executeAgentMutation(task, {
      commandId: `w${workerId}-${cycle}`,
      idempotencyKey: `w${workerId}-${cycle}`,
      payload: [{ command: "set", path: "/body/paragraph[1]", props: { text: `w${workerId}-${cycle}` } }]
    });
    await plugin.flushAgentCandidate(task);
    await plugin.verifyAgentCandidate(task);
    await plugin.finalizeAgentTask(task);
    await plugin.acceptCandidate(session.sessionId, task.candidateId);
    await plugin.closeSession(session.sessionId);
    durations.push(Date.now() - t0);
  } catch (error) {
    errors++;
    firstError ??= `${error?.code ?? ""} ${error?.message ?? String(error)}`.slice(0, 200);
    console.error(`w${workerId} cycle ${cycle} error: ${error?.message ?? error}`);
  }
}
const sorted = [...durations].sort((a, b) => a - b);
const pct = (p) => sorted[Math.min(sorted.length - 1, Math.floor((sorted.length * p) / 100))] ?? -1;
await plugin.dispose().catch(() => undefined);
console.log(
  JSON.stringify({
    workerId,
    cycles,
    accepted: durations.length,
    errors,
    firstError,
    p50: pct(50),
    p95: pct(95)
  })
);
