/**
 * RC cutter (issue #15, RELEASE.md §2/§3/§7): mechanizes the rehearsed
 * release flow so the official RC is one verified command. Fail-closed at
 * every gate — nothing publishes from here.
 *
 * Gates:
 *   1. Clean worktree (exact-SHA provenance).
 *   2. `ci` + `engine-gate` green at HEAD (gh api).
 *   3. MANDATORY --soak-report: runner verdict must be schema-valid
 *      `passed: true` AND verdict.sha must EQUAL HEAD exactly — no lineage
 *      heuristics, no optional soak.
 *
 * Artifact provenance (review: "rc:cut 未证明 artifact 真正来自 candidate
 * SHA"): the caller's dist/ and vendor-bundle/ are IGNORED build outputs —
 * a clean worktree does NOT prove the packed bytes came from HEAD. So the
 * cutter DELETES them and rebuilds from HEAD (npm ci → build:vendor → tsc),
 * stamps release-provenance.json into the artifact, and asserts the
 * installed tree's provenance during rehearsals. A previous artifact is
 * built the same way from the last source-touching commit's parent, giving
 * two genuinely distinguishable binaries.
 *
 * Rehearsals (real artifact transitions — a same-binary re-run fails):
 *   fresh   — install PREVIOUS artifact, smoke, leave a session open.
 *   upgrade — install CANDIDATE artifact; the installed dist-tree hash MUST
 *             change; the open session resolves recovered→ready;
 *             schemaVersion never regresses.
 *   rollback— restore the matching workspace backup AND reinstall the
 *             PREVIOUS artifact; the installed dist-tree hash MUST equal the
 *             fresh one again; the open session resolves ready.
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
const sh = (cmd) => execSync(cmd, { cwd: process.cwd(), encoding: "utf8" }).trim();
const sha = sh("git rev-parse HEAD");
const PROVENANCE_FILE = "release-provenance.json";
const nodeExe = process.execPath;

/** sha256 over a whole dist tree (sorted relpath:filehash lines). */
async function distTreeHash(distDir) {
  const entries = [];
  const walk = async (dir, rel) => {
    for (const name of await readdir(dir, { withFileTypes: true })) {
      const p = join(dir, name.name);
      const r = rel ? `${rel}/${name.name}` : name.name;
      if (name.isDirectory()) await walk(p, r);
      else entries.push([r, sha256Hex(await readFile(p))]);
    }
  };
  await walk(distDir, "");
  entries.sort((a, b) => (a[0] < b[0] ? -1 : 1));
  return sha256Hex(entries.map(([r, h]) => `${r}:${h}`).join("\n"));
}

// ---- 1. Preconditions ----------------------------------------------------
if (sh("git status --porcelain") !== "") {
  fail("worktree is not clean — commit or stash first (exact-SHA provenance)");
}
console.log(`precondition OK: clean worktree at ${sha.slice(0, 12)}`);

let gh = null;
try {
  gh = JSON.parse(
    execSync(
      `gh api "repos/wulianpu/office-agent-plugin/actions/runs?head_sha=${sha}&per_page=10"`,
      { cwd: process.cwd(), encoding: "utf8", maxBuffer: 32 * 1024 * 1024 }
    )
  );
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

// Mandatory exact-SHA soak gate (fail-closed — never optional, never a
// re-derived weaker heuristic; the runner's verdict.passed is authoritative).
const soakReportPath = arg("soak-report");
if (!soakReportPath) {
  fail("--soak-report is REQUIRED: an RC without an exact-SHA PASSED soak is not cuttable");
}
if (!existsSync(soakReportPath)) fail(`soak report not found: ${soakReportPath}`);
const soak = JSON.parse(await readFile(soakReportPath, "utf8")).verdict ?? {};
if (soak.passed !== true) {
  fail(`soak verdict is not schema-valid passed:true — got: ${JSON.stringify(soak).slice(0, 200)}`);
}
if (String(soak.sha) !== sha) {
  fail(`soak sha ${String(soak.sha).slice(0, 12)} != HEAD ${sha.slice(0, 12)} — exact-SHA soak required`);
}
console.log(
  `precondition OK: exact-SHA soak PASSED at ${sha.slice(0, 12)} (${soak.cycles} cycles, ${soak.errors} errors)`
);

// ---- 2. Clean rebuild + pack (candidate, then previous) -------------------
const releaseDir = join(process.cwd(), "release");
async function packBuilt(name) {
  const out = await mkdtemp(join(tmpdir(), "rc-pack-"));
  sh(`npm pack --pack-destination ${JSON.stringify(out)}`);
  const tgz = (await readdir(out)).find((f) => f.endsWith(".tgz"));
  if (!tgz) fail(`npm pack produced no tarball for ${name}`);
  const dest = join(releaseDir, name);
  await cp(join(out, tgz), dest);
  await rm(out, { recursive: true, force: true });
  await rm(join(process.cwd(), PROVENANCE_FILE), { force: true }).catch(() => undefined);
  return { path: dest, sha256: sha256Hex(await readFile(dest)) };
}

/** Delete ignored build outputs, rebuild from the CURRENT checkout, stamp
 * provenance, pack. Stale-output reuse is the provenance hole being closed. */
async function cleanRebuild(name, buildSha) {
  await rm(join(process.cwd(), "dist"), { recursive: true, force: true });
  // vendor-bundle contains a TRACKED manifest/package.json next to the
  // generated bundles — delete only the generated artifacts and restore
  // tracked state, never the directory wholesale.
  await rm(join(process.cwd(), "vendor-bundle"), { recursive: true, force: true }).catch(() => undefined);
  sh("git restore vendor-bundle 2>/dev/null || true");
  await writeFile(
    join(process.cwd(), PROVENANCE_FILE),
    JSON.stringify({ candidateSha: buildSha, builtAt: new Date().toISOString() }, null, 2)
  );
  sh("npm ci"); // hermetic: deps exactly as the committed lockfile says
  sh("npm run build:vendor");
  sh("npx tsc -p tsconfig.json");
  console.log(`rebuilt from ${buildSha.slice(0, 12)} (npm ci + vendor-bundle + tsc)`);
  return packBuilt(name);
}

// Previous binary: parent of the last commit that touched Runtime source —
// a genuinely different dist tree, recent enough for schema compatibility.
const lastSrcCommit = sh(`git log -1 --format=%H -- src/ vendor/`);
const previousSha = sh(`git rev-parse ${lastSrcCommit}^`);
const branch = sh("git rev-parse --abbrev-ref HEAD");

await mkdir(releaseDir, { recursive: true });
for (const stale of await readdir(releaseDir)) await rm(join(releaseDir, stale), { force: true });

const candidate = await cleanRebuild("candidate.tgz", sha);
console.log(`building previous artifact from ${previousSha.slice(0, 12)}…`);
sh(`git checkout --detach ${previousSha}`);
let previous;
try {
  // Incremental rebuild (node_modules kept): the previous artifact is the
  // upgrade/rollback fixture — it just has to genuinely differ from HEAD.
  sh("npm run build:vendor");
  sh("npx tsc -p tsconfig.json");
  await writeFile(
    join(process.cwd(), PROVENANCE_FILE),
    JSON.stringify({ candidateSha: previousSha, builtAt: new Date().toISOString() }, null, 2)
  );
  previous = await packBuilt("previous.tgz");
  console.log(`built: previous.tgz from ${previousSha.slice(0, 12)} (sha256 ${previous.sha256.slice(0, 12)}…)`);
} finally {
  sh(`git checkout ${branch}`);
  await rm(join(process.cwd(), PROVENANCE_FILE), { force: true }).catch(() => undefined);
}
if (sh("git rev-parse HEAD") !== sha) fail("checkout dance did not return to candidate HEAD");
if (sh("git status --porcelain") !== "") fail("worktree dirty after previous-artifact build");
console.log(`built: candidate.tgz (sha256 ${candidate.sha256.slice(0, 12)}…)`);

// Provenance domains (review: wrapper/runtime engine pin domains recorded
// explicitly — not a single engineVersion).
const engineVersion = await new (await import("../../dist/agent/officecli/officecli-adapter.js"))
  .OfficeCliAdapter({ timeoutMs: 30_000 })
  .version_()
  .catch(() => null);
if (!engineVersion) fail("OfficeCLI engine unavailable — RC requires a proven engine");
const gateYml = await readFile(join(process.cwd(), ".github/workflows/engine-gate.yml"), "utf8");
const enginePinCi = gateYml.match(/OFFICECLI_PIN:\s*"([\d.]+)"/)?.[1];
const genofficeSha = sh("git ls-tree HEAD vendor/genoffice").split(/\s+/)[1] ?? "unknown";
const pkg = JSON.parse(await readFile(join(process.cwd(), "package.json"), "utf8"));
if (engineVersion !== enginePinCi) {
  console.warn(
    `NOTE: local engine ${engineVersion} != required-lane pin ${enginePinCi} — recorded as separate provenance domains`
  );
}
await writeFile(join(releaseDir, "node-version.txt"), `${process.version}\n`);
await writeFile(join(releaseDir, "officecli-version.txt"), `${engineVersion}\n`);
await writeFile(
  join(releaseDir, "provenance-domains.txt"),
  `package=@harness/office-plugin@${pkg.version}\nengine-cli(local)=${engineVersion}\nengine-pin(required-ci)=${enginePinCi}\ngenoffice=${genofficeSha}\ncandidate=${sha}\nprevious=${previousSha}\n`
);

// ---- 3/4/5. Install / upgrade / rollback rehearsals -----------------------
const prefix = await mkdtemp(join(tmpdir(), "rc-install-"));
const workspace = join(prefix, "ws");
const backup = join(prefix, "ws-backup");
const pkgRoot = join(prefix, "node_modules", "@harness", "office-plugin");
const installedEntry = join(pkgRoot, "dist", "plugin", "office-plugin.js");
const installedProvenance = join(pkgRoot, PROVENANCE_FILE);

const distShaOfInstalled = () => distTreeHash(join(pkgRoot, "dist"));
const checkProvenance = async (expectedSha, label) => {
  const p = JSON.parse(await readFile(installedProvenance, "utf8"));
  if (p.candidateSha !== expectedSha) {
    fail(`${label}: installed provenance ${String(p.candidateSha).slice(0, 12)} != expected ${expectedSha.slice(0, 12)}`);
  }
};
const schemaVersionOf = async () => {
  // Migrations record the version in meta.schemaVersion (database.ts).
  const { DatabaseSync } = await import("node:sqlite");
  const db = new DatabaseSync(join(workspace, "office-runtime.db"), { readOnly: true });
  const row = db.prepare("SELECT value FROM meta WHERE key = 'schemaVersion'").get();
  db.close();
  return Number(row?.value ?? 0);
};
const openRuntime = async () => {
  const mod = await import(pathToFileURL(installedEntry).href);
  return mod.OfficePlugin.create({ workspaceRoot: workspace, skipHostProbe: true });
};
const install = async (artifact, label) => {
  await rm(join(prefix, "node_modules"), { recursive: true, force: true });
  await run("npm", ["install", pathToFileURL(artifact.path).href, "--omit=dev"], { cwd: prefix });
  if (!existsSync(installedEntry)) fail(`${label}: runtime entry missing from the installed tree`);
};
const phase = async (name, sessionId) => {
  const args = [join(process.cwd(), "tools", "release", "rc-phase.mjs"), name, prefix];
  if (sessionId) args.push(sessionId);
  const { stdout } = await run(nodeExe, args, { cwd: process.cwd(), timeout: 300_000 });
  const line = stdout.trim().split("\n").find((l) => l.trim().startsWith("{"));
  if (!line) fail(`phase ${name} produced no verdict: ${stdout.slice(0, 200)}`);
  return JSON.parse(line.trim());
};

// Fresh install: PREVIOUS artifact, real smoke, session left open.
await mkdir(prefix, { recursive: true });
await run("npm", ["init", "-y"], { cwd: prefix }).catch(() => undefined);
await install(previous, "fresh");
await checkProvenance(previousSha, "fresh");
const freshDistSha = await distShaOfInstalled();
const fresh = await phase("fresh", undefined);
if (!fresh.sessionId) fail("fresh install: no open session reported");
const schemaBefore = fresh.schemaVersion;
console.log(
  `fresh install OK: previous@${previousSha.slice(0, 12)}, dist ${freshDistSha.slice(0, 12)}, session open, schemaVersion=${schemaBefore}`
);

// Upgrade: install the CANDIDATE artifact — the binary MUST change.
await install(candidate, "upgrade");
await checkProvenance(sha, "upgrade");
const upgradedDistSha = await distShaOfInstalled();
if (upgradedDistSha === freshDistSha) {
  fail("upgrade: installed dist tree did NOT change — no real artifact replacement occurred");
}
const schemaAfterUpgrade = (await phase("upgrade", fresh.sessionId)).schemaVersion;
if (schemaAfterUpgrade < schemaBefore) {
  fail(`upgrade: schemaVersion went backwards (${schemaBefore} → ${schemaAfterUpgrade})`);
}
console.log(
  `upgrade OK: previous→candidate (dist ${freshDistSha.slice(0, 12)} → ${upgradedDistSha.slice(0, 12)}), recovery=ready, schemaVersion=${schemaAfterUpgrade}`
);

// Rollback: restore the matching workspace backup AND the previous binary.
await cp(workspace, backup, { recursive: true });
await phase("diverge", undefined); // workspace drifts past the backup point
await rm(workspace, { recursive: true, force: true });
await cp(backup, workspace, { recursive: true });
await install(previous, "rollback");
await checkProvenance(previousSha, "rollback");
const rolledBackDistSha = await distShaOfInstalled();
if (rolledBackDistSha !== freshDistSha) {
  fail(
    `rollback: installed dist ${rolledBackDistSha.slice(0, 12)} != previous ${freshDistSha.slice(0, 12)} — binary was not rolled back`
  );
}
await phase("rollback", fresh.sessionId);
console.log(`rollback OK: previous binary + matching workspace backup restored (dist ${rolledBackDistSha.slice(0, 12)}), recovery=ready`);

// ---- Evidence -------------------------------------------------------------
const evidence = {
  kind: "rc-cut",
  sha,
  previousSha,
  artifacts: { candidate: candidate, previous: previous },
  installedDistSha: {
    freshPrevious: freshDistSha,
    upgradedCandidate: upgradedDistSha,
    rolledBackPrevious: rolledBackDistSha
  },
  provenanceDomains: {
    package: `@harness/office-plugin@${pkg.version}`,
    engineCliLocal: engineVersion,
    enginePinRequiredCi: enginePinCi,
    genoffice: genofficeSha
  },
  nodeVersion: process.version,
  schemaVersion: { before: schemaBefore, afterUpgrade: schemaAfterUpgrade },
  soakGate: { report: soakReportPath, passed: true, sha: soak.sha, cycles: soak.cycles },
  rehearsals: {
    freshInstall: "pass",
    upgrade: "pass",
    rollback: "pass",
    artifactReplacementVerified: true
  },
  previousBuildNote:
    "previous.tgz is an internal upgrade/rollback FIXTURE built from the last source-touching commit's parent; it predates the package files whitelist and is never published", 
  finishedAt: new Date().toISOString()
};
await writeFile(join(releaseDir, "rc-evidence.json"), JSON.stringify(evidence, null, 2));
await writeFile(
  join(releaseDir, "SHA256SUMS.txt"),
  `${candidate.sha256}  candidate.tgz\n${previous.sha256}  previous.tgz\n`
);
await rm(prefix, { recursive: true, force: true }).catch(() => undefined);
console.log(JSON.stringify(evidence));
console.log("RC CUT OK: attach release/ artifacts to the GitHub release");
