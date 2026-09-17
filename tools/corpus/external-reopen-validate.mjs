/**
 * Issue #15 external-validation upgrade: for every production corpus file,
 * drive the REAL Runtime Agent write path (mutate with a unique marker →
 * flush → verify → finalize → accept), then open the SAVED OUTPUT in the
 * real WPS suite application (KWPS/KET/KWPP COM) and assert the marker is
 * readable there. This is the actual "Office/WPS reopens the output"
 * acceptance — independent of the Runtime AND of the OfficeCLI engine, and
 * impossible to satisfy with a ZIP smoke.
 *
 * The engine AND the WPS suite are both REQUIRED here: either missing is a
 * FAIL (never a silent skip).
 *
 * Evidence JSON (candidate SHA, output hashes, per-file verdict) lands in
 * test-results/external-reopen-evidence-<sha12>.json.
 *
 * Usage: node tools/corpus/external-reopen-validate.mjs
 */

import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { copyFile, mkdtemp, readFile, rm, writeFile, mkdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { execSync } from "node:child_process";
import { createHash } from "node:crypto";
import { OfficePlugin } from "../../dist/plugin/office-plugin.js";
import { OfficeCliAdapter } from "../../dist/agent/officecli/officecli-adapter.js";

const run = promisify(execFile);
const sha256Hex = (data) => createHash("sha256").update(data).digest("hex");
const sha = execSync("git rev-parse HEAD", { cwd: process.cwd() }).toString().trim();

const CORPUS_DIR = join(process.cwd(), "corpus", "production");
const MANIFEST_PATH = join(CORPUS_DIR, "manifest.json");

// Both validators are REQUIRED — fail closed, never skip.
const engineAdapter = new OfficeCliAdapter({ timeoutMs: 90_000 });
const engineVersion = await engineAdapter.version_().catch(() => null);
if (!engineVersion) {
  console.error("EXTERNAL REOPEN FAIL: OfficeCLI engine unavailable");
  process.exit(1);
}
try {
  await run("powershell", [
    "-NoProfile", "-Command",
    "$app = New-Object -ComObject KWPS.Application; $app.Quit()"
  ], { timeout: 120_000 });
} catch {
  console.error("EXTERNAL REOPEN FAIL: WPS suite (KWPS COM) unavailable");
  process.exit(1);
}

function mutationTargetFor(file, outline) {
  const ext = file.slice(file.lastIndexOf(".")).toLowerCase();
  if (ext === ".docx") return { path: "/body/paragraph[1]", props: { text: null } };
  if (ext === ".pptx") return { path: "/slide[1]/shape[1]", props: { text: null } };
  if (ext === ".xlsx") {
    const sheets = outline.sheets ?? [];
    if (outline.kind !== "xlsx" || sheets.length === 0) {
      throw new Error(`xlsx corpus ${file} has no sheet in the preview outline`);
    }
    return { path: `/${sheets[0].name}/A1`, props: { value: null } };
  }
  throw new Error(`no mutation mapping for ${file}`);
}

const workspace = await mkdtemp(join(tmpdir(), "external-reopen-"));
const plugin = await OfficePlugin.create({ workspaceRoot: join(workspace, "rt"), skipHostProbe: true });
const manifest = JSON.parse(await readFile(MANIFEST_PATH, "utf8"));
const startedAt = new Date().toISOString();
const results = [];
let failures = 0;

try {
  for (const entry of manifest.files) {
    const marker = `wps-reopen-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const work = join(workspace, entry.file);
    await copyFile(join(CORPUS_DIR, entry.file), work);
    const record = {
      file: entry.file,
      producer: entry.producer,
      inputSha256: sha256Hex(await readFile(work)),
      marker,
      runtimeAccept: false,
      outputSha256: null,
      reopen: { opened: false, markerFound: false, app: null }
    };
    try {
      // ---- Runtime Agent write path commits the marker into the artifact ----
      const ref = await plugin.registerArtifact(work);
      const preview = await plugin.preview({ artifactRef: ref, priority: "visible" });
      const target = mutationTargetFor(entry.file, preview.model.outline);
      const session = await plugin.openSession(ref);
      await plugin.service.sessions.ensureStrongIdentity(session.sessionId);
      const task = await plugin.beginAgentTask(session.sessionId, {
        intent: `external reopen validation (${entry.file})`,
        destructiveAllowed: false
      });
      const props = target.props.text === null ? { text: marker } : { value: marker };
      await plugin.executeAgentMutation(task, {
        commandId: `reopen-${Date.now()}`,
        idempotencyKey: `reopen-${Date.now()}`,
        payload: [{ command: "set", path: target.path, props }]
      });
      await plugin.flushAgentCandidate(task);
      const verification = await plugin.verifyAgentCandidate(task);
      if (verification.structural.status === "fail") {
        throw new Error(`structural verify failed for ${entry.file}`);
      }
      await plugin.finalizeAgentTask(task);
      await plugin.acceptCandidate(session.sessionId, task.candidateId);
      await plugin.closeSession(session.sessionId);
      record.runtimeAccept = true;
      record.outputSha256 = sha256Hex(await readFile(work));
      if (record.outputSha256 === record.inputSha256) {
        throw new Error(`accept did not mutate ${entry.file}`);
      }

      // ---- REAL WPS application reopens the saved output ----
      const stdout = await run("powershell", [
        "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", join(process.cwd(), "tools", "corpus", "external-reopen.ps1"),
        "-File", work, "-Marker", marker
      ], { timeout: 180_000, encoding: "utf8" }).then(r => r.stdout).catch(e => e.stdout ?? "");
      const line = stdout.split("\n").find((l) => l.trim().startsWith("{"));
      if (!line) throw new Error(`reopen helper produced no verdict: ${stdout.slice(0, 200)}`);
      const verdict = JSON.parse(line.trim());
      record.reopen = verdict;
    } catch (error) {
      failures++;
      record.error = String(error?.message ?? error).slice(0, 300);
      console.error(`${entry.file}: ${record.error}`);
    }
    results.push(record);
    const ok = record.reopen.opened && record.reopen.markerFound;
    console.log(`${ok ? "PASS" : "FAIL"} ${entry.file} (app=${record.reopen.app ?? "?"})`);
  }
} finally {
  await plugin.dispose().catch(() => undefined);
  await rm(workspace, { recursive: true, force: true }).catch(() => undefined);
}

const evidence = {
  kind: "external-reopen",
  sha,
  engineVersion,
  startedAt,
  finishedAt: new Date().toISOString(),
  failures,
  files: results
};
await mkdir(join(process.cwd(), "test-results"), { recursive: true });
const evidencePath = join(process.cwd(), "test-results", `external-reopen-evidence-${sha.slice(0, 12)}.json`);
await writeFile(evidencePath, JSON.stringify(evidence, null, 2));
console.log(JSON.stringify({ evidencePath, failures, files: results.length }));
console.log(failures === 0 ? "EXTERNAL REOPEN PASS: WPS suite read every Runtime output" : "EXTERNAL REOPEN FAIL");
process.exit(failures === 0 ? 0 : 1);
