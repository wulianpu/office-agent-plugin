/**
 * One runtime phase of the RC rehearsal, run as a CHILD process so the
 * installed artifact's native modules (sharp/libvips) never lock inside the
 * cutter process — the upgrade step must be able to fully replace the
 * install, and each phase is a genuine cold start (no shared state).
 *
 * Usage: node tools/release/rc-phase.mjs <phase> <prefix> [sessionId]
 *   fresh    — smoke register→preview→open, LEAVE the session open, dispose
 *   upgrade  — resolve the given session recovered→ready, report schemaVersion
 *   diverge  — add + close a second session (workspace drifts past backup)
 *   rollback — resolve the given session recovered→ready after restore
 * Prints one JSON line: {"phase":…, "sessionId":…, "schemaVersion":…}
 */

import { join } from "node:path";
import { pathToFileURL } from "node:url";
import { DatabaseSync } from "node:sqlite";

const [phase, prefix, sessionId] = process.argv.slice(2);
const fail = (msg) => {
  console.error(`PHASE ${phase} FAIL: ${msg}`);
  process.exit(1);
};
const workspace = join(prefix, "ws");
const pkgRoot = join(prefix, "node_modules", "@harness", "office-plugin");
const installedEntry = join(pkgRoot, "dist", "plugin", "office-plugin.js");

const schemaVersionOf = () => {
  // Migrations record the version in meta.schemaVersion (database.ts), not
  // PRAGMA user_version.
  const db = new DatabaseSync(join(workspace, "office-runtime.db"), { readOnly: true });
  const row = db.prepare("SELECT value FROM meta WHERE key = 'schemaVersion'").get();
  db.close();
  return Number(row?.value ?? 0);
};

const mod = await import(pathToFileURL(installedEntry).href);
const adapterMod = await import(
  pathToFileURL(join(pkgRoot, "dist", "agent", "officecli", "officecli-adapter.js")).href
);
const adapter = new adapterMod.OfficeCliAdapter({ timeoutMs: 90_000 });
const plugin = await mod.OfficePlugin.create({ workspaceRoot: workspace, skipHostProbe: true });

const makeDocx = async (name) => {
  const docx = join(prefix, name);
  await adapter.run(["create", docx, "--json"]).catch(() => undefined);
  await adapter.close(docx).catch(() => undefined);
  return docx;
};

try {
  if (phase === "fresh") {
    const docx = await makeDocx("smoke.docx");
    const ref = await plugin.registerArtifact(docx);
    await plugin.preview({ artifactRef: ref, priority: "visible" });
    const session = await plugin.openSession(ref);
    await plugin.service.sessions.ensureStrongIdentity(session.sessionId);
    // No closeSession — dispose with the session open (crash-like stop);
    // this open session is the recovery subject for upgrade/rollback.
    console.log(
      JSON.stringify({ phase, sessionId: session.sessionId, schemaVersion: schemaVersionOf() })
    );
  } else if (phase === "upgrade" || phase === "rollback") {
    const outcome = await plugin.service.resolveRecoveredSession(sessionId);
    if (outcome !== "ready") fail(`recovered session resolved to ${outcome}, expected ready`);
    console.log(JSON.stringify({ phase, sessionId, schemaVersion: schemaVersionOf() }));
  } else if (phase === "diverge") {
    const docx2 = await makeDocx("diverge.docx");
    const ref = await plugin.registerArtifact(docx2);
    const s = await plugin.openSession(ref);
    await plugin.service.sessions.ensureStrongIdentity(s.sessionId);
    await plugin.closeSession(s.sessionId);
    console.log(JSON.stringify({ phase, schemaVersion: schemaVersionOf() }));
  } else {
    fail(`unknown phase ${phase}`);
  }
} finally {
  await plugin.dispose().catch(() => undefined);
}
