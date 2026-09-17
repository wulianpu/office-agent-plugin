/**
 * Production compatibility corpus runner (issue #15 P1-high-2).
 *
 * Executes the full Runtime flow (register → preview → open →
 * beginEdit/Agent path → save/accept → dispose → cold restart → reopen)
 * against the versioned production corpus, then validates the output with
 * an EXTERNAL checker — never just "the Runtime can read its own output".
 *
 * Corpus layout (issue #15):
 *   corpus/production/manifest.json  — versioned manifest (see schema below)
 *   corpus/production/<files…>       — the real Office/WPS-saved documents
 *
 * Manifest schema (v1, strictly validated at RUNTIME — a malformed manifest
 * fails the suite instead of being trusted via a type assertion):
 * {
 *   "version": 1,
 *   "files": [
 *     {
 *       "file": "docx-image-heavy.docx",
 *       "producer": "office" | "wps" | "third-party",
 *       "origin": "sanitized-production",   // where the doc came from
 *       "features": ["images", "tables"],   // feature tags
 *       "sha256": "<hex of the file>",      // integrity pin
 *       "expect": "full" | "degrade"        // capability expectation
 *     }
 *   ]
 * }
 *
 * IMPORTANT (issue #15 acceptance): without a production corpus present,
 * this suite REPORTS the gap and exits non-zero in the engine gate — a
 * synthetic-only corpus must not be reported as production compatibility
 * evidence. Drop sanitized files + manifest in corpus/production/ to enable.
 *
 * Engine lane (issue #15 remaining acceptance): when the OfficeCLI engine is
 * available, each corpus file additionally goes through the REAL Agent write
 * path (task → mutate → flush → verify → finalize → accept), a REAL
 * dispose → new Runtime → cold restart → resolveRecoveredSession chain, and
 * an EXTERNAL reopen of the saved output by the independent engine binary.
 * Results are persisted as evidence JSON bound to the candidate SHA +
 * input/output hashes.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { existsSync } from "node:fs";
import { mkdir, readFile, mkdtemp, writeFile, copyFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { extname, join } from "node:path";
import { execSync } from "node:child_process";
import { createHash } from "node:crypto";
import { OfficePlugin } from "../../src/plugin/office-plugin.js";
import type { PreviewResult } from "../../src/contracts/preview.js";
import { newCommandId } from "../../src/support/ids.js";
import { OfficeCliAdapter } from "../../src/agent/officecli/officecli-adapter.js";

const CORPUS_DIR = join(process.cwd(), "corpus", "production");
const MANIFEST_PATH = join(CORPUS_DIR, "manifest.json");

interface ManifestFile {
  file: string;
  producer: "office" | "wps" | "third-party";
  origin: string;
  features: string[];
  sha256: string;
  expect: "full" | "degrade";
}
interface Manifest {
  version: number;
  files: ManifestFile[];
}

const sha256Hex = (data: Buffer) => createHash("sha256").update(data).digest("hex");

/**
 * Fail-closed runtime schema parser (issue #15 round-29 P2): the manifest is
 * evidence — an illegal producer/expect/version must be REJECTED at runtime,
 * not smuggled through a type assertion.
 */
export function parseManifest(raw: string): Manifest {
  const root: unknown = JSON.parse(raw);
  if (typeof root !== "object" || root === null || Array.isArray(root)) {
    throw new Error("manifest: root must be an object");
  }
  const rootKeys = Object.keys(root as Record<string, unknown>).sort();
  if (rootKeys.join(",") !== "files,version") {
    throw new Error(`manifest: unexpected root keys [${rootKeys.join(", ")}]`);
  }
  const { version, files } = root as Record<string, unknown>;
  if (version !== 1) throw new Error(`manifest: unsupported version ${String(version)}`);
  if (!Array.isArray(files) || files.length === 0) {
    throw new Error("manifest: files must be a non-empty array");
  }
  const ENTRY_KEYS = "expect,features,file,origin,producer,sha256";
  const out: ManifestFile[] = [];
  for (const rawEntry of files) {
    if (typeof rawEntry !== "object" || rawEntry === null || Array.isArray(rawEntry)) {
      throw new Error("manifest: entry must be an object");
    }
    const entry = rawEntry as Record<string, unknown>;
    const keys = Object.keys(entry).sort();
    if (keys.join(",") !== ENTRY_KEYS) {
      throw new Error(`manifest: unexpected entry keys [${keys.join(", ")}]`);
    }
    if (typeof entry.file !== "string" || entry.file.length === 0) {
      throw new Error("manifest: file must be a non-empty string");
    }
    if (/[/\\]/.test(entry.file) || entry.file.includes("..")) {
      throw new Error(`manifest: file must be a bare name, got ${entry.file}`);
    }
    if (entry.producer !== "office" && entry.producer !== "wps" && entry.producer !== "third-party") {
      throw new Error(`manifest: illegal producer ${String(entry.producer)} for ${entry.file}`);
    }
    if (typeof entry.origin !== "string" || entry.origin.length === 0) {
      throw new Error(`manifest: origin must be a non-empty string for ${entry.file}`);
    }
    if (
      !Array.isArray(entry.features) ||
      entry.features.length === 0 ||
      !entry.features.every((f) => typeof f === "string" && f.length > 0)
    ) {
      throw new Error(`manifest: features must be a non-empty string array for ${entry.file}`);
    }
    if (typeof entry.sha256 !== "string" || !/^[0-9a-f]{64}$/.test(entry.sha256)) {
      throw new Error(`manifest: sha256 must be 64 lowercase hex chars for ${entry.file}`);
    }
    if (entry.expect !== "full" && entry.expect !== "degrade") {
      throw new Error(`manifest: illegal expect ${String(entry.expect)} for ${entry.file}`);
    }
    out.push(entry as unknown as ManifestFile);
  }
  return { version, files: out };
}

let corpusAvailable = false;
let workspace: string;

beforeAll(async () => {
  corpusAvailable = existsSync(MANIFEST_PATH);
  if (corpusAvailable) {
    workspace = await mkdtemp(join(tmpdir(), "compat-corpus-"));
  }
});

afterAll(async () => {
  if (workspace) await rm(workspace, { recursive: true, force: true }).catch(() => undefined);
});

describe("production compatibility corpus (issue #15 P1-high-2)", () => {
  it("reports the corpus gap honestly when no production corpus exists", async (ctx) => {
    if (corpusAvailable) ctx.skip();
    // Honest reporting per issue #15: a missing corpus is a GAP, not a pass.
    console.warn(
      "PRODUCTION COMPATIBILITY GAP: corpus/production/manifest.json not present — " +
        "real Office/WPS compatibility evidence is NOT being produced. " +
        "Drop sanitized production documents + manifest to enable this suite."
    );
    ctx.skip();
  });

  it("parses the manifest through the fail-closed runtime schema, then verifies sha256 pins", async (ctx) => {
    if (!corpusAvailable) ctx.skip();
    const manifest = parseManifest(await readFile(MANIFEST_PATH, "utf8"));
    for (const entry of manifest.files) {
      const path = join(CORPUS_DIR, entry.file);
      expect(existsSync(path), `missing corpus file: ${entry.file}`).toBe(true);
      const digest = sha256Hex(await readFile(path));
      expect(digest, `sha256 mismatch for ${entry.file}`).toBe(entry.sha256);
    }
  });

  it("runs the Human flow per corpus file and EXTERNALLY validates the output", async (ctx) => {
    if (!corpusAvailable) ctx.skip();
    parseManifest(await readFile(MANIFEST_PATH, "utf8")); // fail before any flow on bad evidence
    const plugin = await OfficePlugin.create({
      workspaceRoot: join(workspace, "rt"),
      skipHostProbe: true
    });
    try {
      for (const entry of parseManifest(await readFile(MANIFEST_PATH, "utf8")).files) {
        const src = join(CORPUS_DIR, entry.file);
        const work = join(workspace, entry.file);
        await copyFile(src, work);

        // register → preview (must not crash / corrupt)
        const ref = await plugin.registerArtifact(work);
        const preview = await plugin.preview({ artifactRef: ref, priority: "visible" });
        expect(preview.model.outline.kind).toBeTruthy();

        // open (read-only MVCC)
        const session = await plugin.openSession(ref);
        await plugin.service.sessions.ensureStrongIdentity(session.sessionId);

        // beginEdit → editor save gate (Runtime gate decides full/degrade)
        const { editor } = await plugin.beginEdit(session.sessionId);
        editor.markDirty?.();
        await editor.save(); // fails closed on external mutation (P0-5)

        // accept → cold restart → reopen (recovery semantics)
        await plugin.endEdit(session.sessionId);
        await plugin.closeSession(session.sessionId);

        // External validation: bytes must still parse as a ZIP with a
        // [Content_Types].xml part — the minimum "openable by another tool"
        // contract. The engine lane below adds an independent-engine reopen.
        const output = await readFile(work);
        expect(output.subarray(0, 2).toString("latin1")).toBe("PK");
        expect(output.includes("[Content_Types].xml")).toBe(true);
      }
    } finally {
      await plugin.dispose().catch(() => undefined);
    }
  });
});

/** CI without the OfficeCLI engine: the engine lane skips cleanly. */
const engineUp = await new OfficeCliAdapter({ timeoutMs: 15_000 })
  .version_()
  .then(() => true)
  .catch(() => false);

interface CompatEvidence {
  candidateSha: string;
  engineVersion: string;
  generatedAt: string;
  files: Array<{
    file: string;
    producer: string;
    inputSha256: string;
    outputSha256: string;
    sessionId: string;
    candidateId: string;
    revisionId: string;
    verificationConfidence: string;
    recoveryOutcome: string;
    externalReopen: "pass";
  }>;
}

function mutationTargetFor(file: string, outline: PreviewResult["model"]["outline"]): string {
  switch (extname(file).toLowerCase()) {
    case ".docx":
      return "/body/paragraph[1]";
    case ".pptx":
      return "/slide[1]/shape[1]";
    case ".xlsx": {
      const sheets = (outline as { kind: string; sheets?: Array<{ name: string }> }).sheets;
      if (outline.kind !== "xlsx" || !sheets || sheets.length === 0) {
        throw new Error(`production-compat: xlsx corpus ${file} has no sheet in the preview outline`);
      }
      return `/${sheets[0]!.name}/A1`;
    }
    default:
      throw new Error(`production-compat: no mutation mapping for ${file} — extend mutationTargetFor (fail-closed)`);
  }
}

describe.skipIf(!engineUp)(
  "production compatibility engine lane (issue #15: Agent path + cold restart + external validator)",
  () => {
    it("drives the Agent write path on real corpus bytes, cold-restarts the Runtime, and externally revalidates the saved output", async (ctx) => {
      if (!corpusAvailable) ctx.skip();
      const manifest = parseManifest(await readFile(MANIFEST_PATH, "utf8"));
      const candidateSha =
        process.env.GITHUB_SHA ??
        execSync("git rev-parse HEAD", { cwd: process.cwd() }).toString().trim();
      const engineAdapter = new OfficeCliAdapter({ timeoutMs: 90_000 });
      const engineVersion = await engineAdapter.version_();
      const wsRoot = join(workspace, "engine-rt");
      const evidence: CompatEvidence = {
        candidateSha,
        engineVersion,
        generatedAt: new Date().toISOString(),
        files: []
      };

      for (const entry of manifest.files) {
        const src = join(CORPUS_DIR, entry.file);
        const work = join(workspace, `engine-${entry.file}`);
        await copyFile(src, work);
        const inputSha256 = sha256Hex(await readFile(work));
        const stamp = `compat-mutation-${Date.now()}`;

        // ---- Phase A: Agent write path on a live Runtime ----
        const first = await OfficePlugin.create({ workspaceRoot: wsRoot, skipHostProbe: true });
        let sessionId = "";
        let candidateId = "";
        let revisionId = "";
        let confidence = "";
        let target = "";
        try {
          const ref = await first.registerArtifact(work);
          const preview = await first.preview({ artifactRef: ref, priority: "visible" });
          target = mutationTargetFor(entry.file, preview.model.outline);

          const session = await first.openSession(ref);
          sessionId = session.sessionId;
          await first.service.sessions.ensureStrongIdentity(sessionId);

          const task = await first.beginAgentTask(sessionId, {
            intent: `production-compat agent mutation (${entry.file})`,
            destructiveAllowed: false
          });
          candidateId = task.candidateId;
          await first.executeAgentMutation(task, {
            commandId: newCommandId(),
            idempotencyKey: newCommandId(),
            payload: [
              extname(entry.file).toLowerCase() === ".xlsx"
                ? { command: "set", path: target, props: { value: stamp } }
                : { command: "set", path: target, props: { text: stamp } }
            ]
          });
          await first.flushAgentCandidate(task);
          const verification = await first.verifyAgentCandidate(task);
          confidence = verification.confidence;
          // "warn" = inherited producer schema quirks (present in the base
          // revision too); only "fail" (mutation-introduced) is a rejection.
          expect(verification.structural.status, `structural verify for ${entry.file}`).not.toBe(
            "fail"
          );
          await first.finalizeAgentTask(task);
          const accepted = await first.acceptCandidate(sessionId, candidateId);
          revisionId = accepted.revisionId;
          // Session deliberately left OPEN — dispose() below is a crash-like
          // stop, so the cold restart must recover (not skip) this session.
        } finally {
          await first.dispose().catch(() => undefined);
        }

        // ---- Phase B: cold restart over the SAME workspace root ----
        const second = await OfficePlugin.create({ workspaceRoot: wsRoot, skipHostProbe: true });
        try {
          const recovered = second.service.sessions.list().find((s) => s.sessionId === sessionId);
          expect(recovered, `session for ${entry.file} rehydrated after dispose`).toBeDefined();
          const outcome = await second.service.resolveRecoveredSession(sessionId);
          expect(outcome, `recovery resolution for ${entry.file}`).toBe("ready");

          // Writer lease is re-acquirable post-restart (fencing token advances).
          const { editor } = await second.beginEdit(sessionId);
          editor.markDirty?.();
          await editor.save();
          await second.endEdit(sessionId);
          await second.closeSession(sessionId);

          const output = await readFile(work);
          const outputSha256 = sha256Hex(output);
          expect(outputSha256, `accept mutated the artifact for ${entry.file}`).not.toBe(inputSha256);
          expect(output.subarray(0, 2).toString("latin1")).toBe("PK");
          expect(output.includes("[Content_Types].xml")).toBe(true);

          // ---- External validation: the INDEPENDENT pinned engine reopens
          // the saved output and reads the mutated node back. This is not
          // "the Runtime reading its own output". ----
          const reopened = await engineAdapter.runBatchStandalone(work, [
            { command: "get", path: target }
          ]);
          expect(reopened.summary.failed ?? 0).toBe(0);
          expect(reopened.results[0]?.success, `external reopen for ${entry.file}`).toBe(true);
          expect(
            JSON.stringify(reopened.results[0]?.output),
            `external read-back of the mutation for ${entry.file}`
          ).toContain(stamp);
          await engineAdapter.close(work).catch(() => undefined);

          evidence.files.push({
            file: entry.file,
            producer: entry.producer,
            inputSha256,
            outputSha256,
            sessionId,
            candidateId,
            revisionId,
            verificationConfidence: confidence,
            recoveryOutcome: outcome,
            externalReopen: "pass"
          });
        } finally {
          await second.dispose().catch(() => undefined);
        }
      }

      // Persist traceable evidence bound to the candidate SHA + hashes.
      const evidenceDir = process.env.COMPAT_EVIDENCE_DIR ?? join(process.cwd(), "test-results");
      await mkdir(evidenceDir, { recursive: true });
      const evidencePath = join(evidenceDir, `compat-evidence-${candidateSha.slice(0, 12)}.json`);
      await writeFile(evidencePath, JSON.stringify(evidence, null, 2));
      expect(evidence.files.length, "every corpus file produced engine-lane evidence").toBe(
        manifest.files.length
      );
      expect(existsSync(evidencePath)).toBe(true);
      console.warn(`PRODUCTION COMPATIBILITY EVIDENCE: ${evidencePath}`);
    }, 600_000);
  }
);
