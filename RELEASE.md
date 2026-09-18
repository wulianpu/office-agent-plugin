# Release & Operations (issue #15 P1)

This document is the Release/Operations contract for the Office Runtime
Plugin: how a release is produced, what evidence must accompany it, and how
to operate/diagnose/roll back in production.

## 1. Release evidence contract

A release candidate is identified by one immutable **candidate SHA**. Before
announcing production readiness, the following evidence must exist and be
linked to that SHA:

| Evidence | Where | Enforced by |
| --- | --- | --- |
| Windows + Linux correctness CI green | `ci.yml` required checks | Ruleset `main-required-ci-gate` (PR-only, strict) |
| Engine-backed lane: pinned OfficeCLI installed/probed, Agent suites executed (not skipped) | `engine-gate.yml` artifact `engine-gate-results` | engine-gate job (any skip = red) |
| OfficeCLI version + compatibility fingerprint | `officecli-version.txt` artifact | engine-gate job |
| Benchmark metrics JSON (TTFP/TTE/TTP p95, heap, RSS) | `bench-metrics.json` artifact (uploaded by the `bench` job) | engine-gate `bench` job + `tools/bench/gate.mjs` hard limits |
| Production compatibility corpus run (when corpus present) | `production-compat.test.ts` results | suite reports the gap honestly if corpus absent |
| Serial engine soak (multi-hour plateau) | `soak-engine-report.json` (`verdict.passed === true`, exact-SHA) | `rc:cut` soak gate (mandatory) |
| Concurrent engine workload (queue wait / OS handles / drain) | `soak-concurrent-report.json` (`verdict.passed === true`) | `tools/soak/run-engine-concurrent.mjs` |
| Benchmark metrics JSON (TTFP/TTE/TTP p95, heap, RSS) | `bench-metrics.json` artifact (uploaded by the `bench` job) | engine-gate `bench` job + `tools/bench/gate.mjs` hard limits |

The pinned OfficeCLI version lives in `engine-gate.yml` (`OFFICECLI_PIN`).
Upgrading it requires a PR that runs the engine-gate against the new version
(compatibility matrix) before the pin moves.

## 1b. Immutability policy

Release tags (`v*`) are protected by the `release-tags-immutable` ruleset
(deletion + non-fast-forward denied, no bypass). Publishing an RC therefore
means: cut a NEW tag (rc.N+1) rather than moving an existing one, and attach
artifacts with their SHA256SUMS recorded inside `rc-evidence.json`. The
RC evidence records provenance domains separately: package version, local
engine CLI, required-CI engine pin, and the GenOffice submodule SHA.

## 2. Producing a release artifact

```text
git checkout <candidate-SHA>
npm ci
npm run build:vendor && npx tsc -p tsconfig.json
# Package the runtime: dist/, vendor-bundle/, package.json, README.md
npm pack --pack-destination release/
# Record integrity + environment:
node -e "console.log(require('crypto').createHash('sha256').update(require('fs').readFileSync('release/harness-office-plugin-3.0.0.tgz')).digest('hex'))" >> release/SHA256SUMS.txt
node --version > release/node-version.txt
officecli --version > release/officecli-version.txt
```

`release/` then contains the immutable artifact, its SHA-256, and the exact
Node/OfficeCLI versions — attach all three to the release record.

The package `files` whitelist enforces the runtime-only contents
(`dist/` + `vendor-bundle/` + `README.md`); a full-repo pack (src/test/corpus
inside the artifact) is a release-blocking defect. Install rehearsal before
any RC announcement: `npm install <tgz> --omit=dev` into a clean prefix, then
import `OfficePlugin` from the installed tree and create/dispose one instance.

`npm run rc:cut` mechanizes §2/§3 end to end and fails closed. Preconditions
(clean worktree; `ci` + `engine-gate` green at HEAD; MANDATORY
`--soak-report <path>` whose runner verdict is schema-valid `passed: true`
and whose `sha` EQUALS HEAD exactly — an RC without an exact-SHA PASSED soak
is not cuttable). The artifact is REBUILT from HEAD, never packed from the
caller's ignored `dist/`/`vendor-bundle/`: those are deleted, then
`npm ci → build:vendor → tsc` run in place, and a `release-provenance.json`
(candidate SHA) is stamped into the package; the install rehearsals assert
the installed provenance matches the expected commit. A previous artifact is
built from the last source-touching commit's parent, giving two genuinely
distinguishable binaries: fresh install uses the previous artifact, upgrade
must observe the installed dist-tree hash CHANGE (else fail), and rollback
restores the previous binary AND the matching workspace backup, with the
dist-tree hash asserted equal again. It writes `release/rc-evidence.json`
(including provenance domains: package, local engine CLI, required-CI engine
pin, GenOffice submodule SHA); it never publishes.

## 3. Install / upgrade / rollback

**Fresh install**: unpack artifact → `npm ci --omit=dev` against it → point
the host at the workspace root → run the smoke flow (`register → preview →
open`) → confirm `officecli --version` matches the matrix.

**Upgrade**: stop the runtime → back up the workspace (see §4) → replace the
artifact → restart. Schema migrations run automatically and are
crash-atomic per version (per-version `BEGIN IMMEDIATE`, see
`src/runtime/persistence/database.ts`); a crash mid-upgrade rolls back and
retries on the next start. Record `schemaVersion` before/after in the
release notes.

**Downgrade/rollback**: restore the previous artifact **and** the backed-up
workspace. The runtime refuses a DB whose `schemaVersion` is NEWER than the
running binary supports (`session DB schema N is newer than runtime supports
M`) — that refusal is the rollback boundary: restore the matching workspace
backup rather than forcing an older binary onto a newer schema.

## 4. Workspace backup & recovery

Everything durable lives under the workspace root:

- `office-runtime.db` (+ WAL/SHM): sessions, revisions, candidates, leases,
  commit journal, idempotency receipts — the recovery source of truth;
- staging/ directories: candidate working copies (safe to delete only when
  no session is active; the eviction ladder purges them otherwise).

**Backup**: copy the whole workspace root while the runtime is stopped
(WAL is checkpointed on clean close). **Restore**: replace the directory and
restart — crash recovery (`RecoveryService`) reconstructs committed state
from journal + filesystem hash facts; interrupted commits resolve to
rolled-back / forward-finalized / conflict by hash truth (INV-15).

## 5. Operations runbook

| Symptom | Diagnosis | Action |
| --- | --- | --- |
| Startup fails: `session DB schema N is newer than runtime supports` | Artifact/workspace version mismatch | Restore the workspace backup matching the binary (see downgrade) |
| OfficeCLI unavailable (`engine unavailabl…` / probe timeout) | Engine not installed / PATH / version drift | Check `officecli --version`; reinstall pinned version. Runtime degrades to hermetic paths (no engine mutations) — agent mutations are refused, previews fall back to the JS renderer |
| Session `recovery-required` after restart | Crash interrupted a session | `resolveRecoveredSession` (or reopen): hash facts decide ready vs conflict. Recovered in-flight candidates are quarantined as `failed` with reason `recovered-after-crash` |
| Session `conflict` | External edit detected (watcher or hash seal) | Source holds newer external bytes — reopen the session; do NOT force-accept over it (commits fail closed by design) |
| Candidate stuck / orphaned staging | Task died mid-flight | Candidates are per-session owned; staging is purged by the eviction ladder + startup purge. Manual: delete `staging/` files when no session is active |
| Post-replace hash mismatch event | Filesystem anomaly or rename race | Journal is at SOURCE_REPLACED; Recovery resolves by hash facts. Treat as `recovery-required`, never retry blindly |
| Memory/RSS growth over hours | Cache pressure ladder should trim | Check governor pressure snapshot (`memoryUsage()`), restart if beyond budget; file an issue with `bench-metrics.json` |

## 6. Known limitations (unsupported ≠ corrupted)

- XLSX has no full editor runtime: beginEdit runs the basic editor on a
  metadata context (degraded by design, see `editorProfileFor`);
- Unsupported Office features (tracked changes, embedded media, etc.) are
  not fully preserved through human edit paths — the compatibility corpus
  (`test/integration/production-compat.test.ts`) records per-feature
  expectations; anything not listed as supported degrades and is reported
  in the preview outline, never silently corrupted;
- Direct `officecli` usage outside the Runtime is unsupported (the Runtime
  owns the resident pool, leases and commit gate).

## 7. Release checklist (bind to the candidate SHA)

1. Required CI (Windows + Linux) green at the SHA — enforced by ruleset;
2. `engine-gate` job green: OfficeCLI pin installed, probe logged, engine
   suites executed > 0;
3. `bench-metrics.json` artifact present and `tools/bench/gate.mjs` PASS;
4. Production corpus run executed (or the gap explicitly accepted in the
   release notes);
5. Artifact + SHA256 + node/officecli versions recorded in the release;
6. Upgrade/rollback notes updated if the DB schema version moved.
