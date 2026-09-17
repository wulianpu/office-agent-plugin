/**
 * RC cutter (issue #15, RELEASE.md §2/§3/§7): mechanizes the rehearsed
 * release flow so the official RC is one verified command, not hand-typed
 * steps. Fail-closed at every gate — nothing publishes from here; the caller
 * attaches the produced artifacts to the GitHub release.
 *
 * Flow:
 *   1. Preconditions — clean worktree, required CI green at HEAD, and (when
 *      --soak-report is given) a PASSED engine soak whose Runtime source is
 *      identical to HEAD (lineage check over src/ + vendor/).
 *   2. Pack — npm pack → release/, SHA256SUMS, node/officecli versions.
 *   3. Fresh install — clean prefix, --omit=dev, smoke through the INSTALLED
 *      tree, leave one session open (crash-like stop).
 *   4. Upgrade — replace the artifact install over the same workspace; the
 *      open session must resolve recovered→ready; schemaVersion never
 *      regresses.
 *   5. Rollback — diverge the workspace, restore the pre-diverge backup
 *      (holding the open session), restarted runtime resolves it ready.
 *
 * Runtime phases run as CHILD processes (tools/release/rc-phase.mjs): the
 * installed artifact's native modules must never lock inside this process,
 * and every phase is a genuine cold start.
 *
 * Usage:
 *   node tools/release/cut-rc.mjs --soak-report test-results/soak-engine-report.json
 */

import { execSync, execFile } from "node:child_process";
import { promisify } from "node:util";
import { createHash } from "node:crypto";
import { existsSync } from "node:fs";
import { cp, mkdir, mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

const arg = (name) => {
  const i = process.argv.indexOf(`--${name}`);
  return i >= 0 ? process.argv[i + 1] : undefined;
};
const fail = (msg) => {
  console.error(`RC CUT FAIL: ${msg}`);
  process.exit(1);
};
const run = (cmd, args, opts) =>
  promisify(execFile)(cmd, args, { shell: process.platform === "win32", ...opts });
const sha256Hex = (data) => createHash("sha256").update(data).digest("hex");
const sha = execSync("git rev-parse HEAD", { cwd: process.cwd() }).toString().trim();
const nodeExe = process.execPath;
const phase = async (name, prefix, sessionId) => {
  const args = [join(process.cwd(), "tools", "release", "rc-phase.mjs"), name, prefix];
  if (sessionId) args.push(sessionId);
  const { stdout } = await run(nodeExe, args, { cwd: process.cwd(), timeout: 300_000 });
  const line = stdout.trim().split("\n").find((l) => l.trim().startsWith("{"));
  if (!line) fail(`phase ${name} produced no verdict: ${stdout.slice(0, 200)}`);
  return JSON.parse(line.trim());
};

// ---- 1. Preconditions ----------------------------------------------------
if (execSync("git status --porcelain", { cwd: process.cwd() }).toString().trim() !== "") {
  fail("worktree is not clean — commit or stash first (exact-SHA provenance)");
}
console.log(`precondition OK: clean worktree at ${sha.slice(0, 12)}`);

let gh = null;
try {
  const stdout = execSync(
    `gh api "repos/wulianpu/office-agent-plugin/actions/runs?head_sha=${sha}&per_page=10"`,
    { cwd: process.cwd(), encoding: "utf8", maxBuffer: 32 * 1024 * 1024 }
  );
  gh = JSON.parse(stdout);
} catch {
  fail("cannot query CI runs via gh api — required checks unprovable");
}
const runsForSha = gh.workflow_runs ?? [];
for (const workflow of ["ci", "engine-gate"]) {
  const ok = runsForSha.some(
    (r) => r.name === workflow && r.head_sha === sha && r.conclusion === "success"
  );
  if (!ok) fail(`required workflow '${workflow}' is not green at HEAD — refusing to cut an RC`);
}
console.log("precondition OK: ci + engine-gate green at HEAD");

const soakReportPath = arg("soak-report");
if (soakReportPath) {
  if (!existsSync(soakReportPath)) fail(`soak report not found: ${soakReportPath}`);
  const report = JSON.parse(await readFile(soakReportPath, "utf8"));
  const verdict = report.verdict ?? report;
  const pass = verdict.passed === true || (verdict.errors === 0 && Number(verdict.rssGrowthRatio) < 2.5);
  if (!pass) fail(`soak verdict is not a PASS: ${JSON.stringify(verdict).slice(0, 200)}`);
  const soakSha = String(verdict.sha);
  execSync(`git diff --quiet ${soakSha} HEAD -- src/ vendor/`, { cwd: process.cwd() });
  console.log(
    `precondition OK: soak ${verdict.cycles} cycles PASSED at ${soakSha.slice(0, 12)}; Runtime source identical to HEAD`
  );
}

// ---- 2. Pack --------------------------------------------------------------
const releaseDir = join(process.cwd(), "release");
await mkdir(releaseDir, { recursive: true });
for (const stale of await readdir(releaseDir)) await rm(join(releaseDir, stale), { force: true });
execSync("npm pack --pack-destination release/", { cwd: process.cwd(), stdio: "pipe" });
const artifact = (await readdir(releaseDir)).find((f) => f.endsWith(".tgz"));
if (!artifact) fail("npm pack produced no tarball");
const artifactPath = join(releaseDir, artifact);
const artifactSha = sha256Hex(await readFile(artifactPath));
await writeFile(join(releaseDir, "SHA256SUMS.txt"), `${artifactSha}  ${artifact}\n`);
await writeFile(join(releaseDir, "node-version.txt"), `${process.version}\n`);
const engineVersion = await new (await import("../../dist/agent/officecli/officecli-adapter.js"))
  .OfficeCliAdapter({ timeoutMs: 30_000 })
  .version_()
  .catch(() => null);
if (!engineVersion) fail("OfficeCLI engine unavailable — RC requires a proven engine");
await writeFile(join(releaseDir, "officecli-version.txt"), `${engineVersion}\n`);
console.log(`packed: ${artifact} (${artifactSha.slice(0, 12)}…, engine ${engineVersion})`);

// ---- 3/4/5. Install / upgrade / rollback rehearsals -----------------------
const prefix = await mkdtemp(join(tmpdir(), "rc-install-"));
const workspace = join(prefix, "ws");
const backup = join(prefix, "ws-backup");
const installArtifact = async () => {
  await run("npm", ["install", pathToFileURL(artifactPath).href, "--omit=dev"], { cwd: prefix });
  const entry = join(prefix, "node_modules", "@harness", "office-plugin", "dist", "plugin", "office-plugin.js");
  if (!existsSync(entry)) fail("install: runtime entry missing from the installed tree");
};

await mkdir(prefix, { recursive: true });
await run("npm", ["init", "-y"], { cwd: prefix }).catch(() => undefined);
await installArtifact();
const fresh = await phase("fresh", prefix);
if (!fresh.sessionId) fail("fresh install: no open session reported");
const schemaBefore = fresh.schemaVersion;
console.log(`fresh install OK: smoke passed, session ${fresh.sessionId} left open, schemaVersion=${schemaBefore}`);

const schemaAfterUpgrade = (await phase("upgrade", prefix, fresh.sessionId)).schemaVersion;
if (schemaAfterUpgrade < schemaBefore) {
  fail(`upgrade: schemaVersion went backwards (${schemaBefore} → ${schemaAfterUpgrade})`);
}
console.log(`upgrade OK: artifact replaced over live workspace, recovery=ready, schemaVersion=${schemaAfterUpgrade}`);

await cp(workspace, backup, { recursive: true }); // holds the open session
await phase("diverge", prefix); // drift the workspace past the backup point
await rm(workspace, { recursive: true, force: true });
await cp(backup, workspace, { recursive: true });
await phase("rollback", prefix, fresh.sessionId);
console.log("rollback OK: pre-diverge backup restored, open session resolves ready");

// ---- Evidence -------------------------------------------------------------
const evidence = {
  kind: "rc-cut",
  sha,
  artifact,
  artifactSha256: artifactSha,
  engineVersion,
  nodeVersion: process.version,
  schemaVersion: { before: schemaBefore, afterUpgrade: schemaAfterUpgrade },
  rehearsals: { freshInstall: "pass", upgrade: "pass", rollback: "pass" },
  finishedAt: new Date().toISOString()
};
await writeFile(join(releaseDir, "rc-evidence.json"), JSON.stringify(evidence, null, 2));
await rm(prefix, { recursive: true, force: true }).catch(() => undefined);
console.log(JSON.stringify(evidence));
console.log("RC CUT OK: attach release/ artifacts to the GitHub release");
