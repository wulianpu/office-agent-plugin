/**
 * Issue #15 round-30: L1 structural validation must NOT reject real-world
 * producer files for schema quirks they inherited from their BASE revision
 * (WPS writes w:uiPriority in a schema-unexpected position in every style).
 * The gate fails closed only against findings the MUTATION introduced, and
 * against infrastructure errors (fail-closed on unresolvable base).
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  VerificationPipeline,
  type VerificationDeps
} from "../../src/verification/pipeline.js";
import {
  findingSignature,
  parseValidationFindings,
  type ValidationFinding,
  type ValidationOutcome
} from "../../src/agent/officecli/officecli-adapter.js";
import { writeDocxFixture } from "../helpers/fixtures.js";

let root: string;
let stagingPath: string;
let livePath: string;

beforeAll(async () => {
  root = await mkdtemp(join(tmpdir(), "l1-quirks-"));
  stagingPath = await writeDocxFixture(join(root, "staging.docx"), ["candidate bytes"]);
  livePath = await writeDocxFixture(join(root, "live.docx"), ["base bytes"]);
});

afterAll(async () => {
  await rm(root, { recursive: true, force: true }).catch(() => undefined);
});

const finding = (schema: string, path?: string, part?: string): ValidationFinding => ({
  schema,
  path,
  part
});

function makeDeps(outcomes: Record<string, ValidationOutcome>): VerificationDeps {
  return {
    // ref "staging" → candidate clone; ref "live" → the base revision bytes.
    store: {
      resolvePath(ref: string) {
        if (ref === "staging") return stagingPath;
        if (ref === "live") return livePath;
        throw new Error(`unresolvable artifactRef: ${ref}`);
      }
    },
    scanner: {
      scan: async () => ({ contentHash: "hash", integrity: { ok: true, corruptEntries: [] } })
    },
    adapter: {
      validate: async (file: string) => {
        const key = file === stagingPath ? "staging" : file === livePath ? "live" : "other";
        const outcome = outcomes[key];
        if (!outcome) return { passed: true, message: "", findings: [] };
        return outcome;
      },
      get: async () => ({ matches: 1 })
    },
    scheduler: {
      submit: (job: { run: () => Promise<unknown> }) => ({ promise: job.run() })
    }
  } as never as VerificationDeps;
}

function makePipeline(outcomes: Record<string, ValidationOutcome>): VerificationPipeline {
  return new VerificationPipeline(makeDeps(outcomes));
}

function verify(pipeline: VerificationPipeline, baseRef: string | undefined) {
  return pipeline.verify({
    candidate: {
      candidateId: "cand_1",
      artifactRef: "staging",
      baseRevisionId: "rev_1"
    } as never,
    baseRevision: { revisionId: "rev_1", artifactRef: baseRef } as never,
    changedTargets: [],
    engineAvailable: true
  });
}

describe("L1 inherited producer quirks (issue #15 round-30)", () => {
  it("parseValidationFindings: engine warning triplets become structured findings", () => {
    const findings = parseValidationFindings([
      { message: "Found 14 validation error(s):" },
      { message: "[Schema] unexpected child element 'uiPriority'." },
      { message: "Path: /w:styles[1]/w:style[2]" },
      { message: "Part: /word/styles.xml" },
      { message: "[Schema] unexpected child element 'uiWrap'." },
      { message: "Path: /w:styles[1]/w:style[3]" },
      { message: "Part: /word/styles.xml" },
      { message: "" },
      { message: "bare engine complaint without structure" }
    ]);
    expect(findings).toHaveLength(3);
    expect(findings[0]).toEqual({
      schema: "[Schema] unexpected child element 'uiPriority'.",
      path: "/w:styles[1]/w:style[2]",
      part: "/word/styles.xml"
    });
    expect(findings[1]?.part).toBe("/word/styles.xml");
    expect(findings[2]?.schema).toBe("bare engine complaint without structure");
    expect(findingSignature(findings[0]!)).not.toBe(findingSignature(findings[1]!));
    expect(findingSignature(findings[0]!)).toBe(findingSignature({ ...findings[0]! }));
  });

  it("findings present identically in the base revision are inherited → warn, not fail", async () => {
    const quirkA = finding("[Schema] uiPriority", "/w:styles[1]/w:style[2]", "/word/styles.xml");
    const quirkB = finding("[Schema] uiPriority", "/w:styles[1]/w:style[5]", "/word/styles.xml");
    const pipeline = makePipeline({
      staging: { passed: false, message: "", findings: [quirkA, quirkB] },
      live: { passed: false, message: "", findings: [quirkA, quirkB] }
    });
    const report = await verify(pipeline, "live");
    expect(report.structural.status).toBe("warn");
    expect(report.structural.issues.every((i) => i.code === "inherited-schema-quirk")).toBe(true);
    // warn does not drag confidence to unverified — the service gate passes.
    expect(report.confidence).toBe("structural");
  });

  it("a finding the base does not have (mutation-introduced) fails closed", async () => {
    const quirkA = finding("[Schema] uiPriority", "/w:styles[1]/w:style[2]", "/word/styles.xml");
    const fresh = finding("[Schema] invalid element 'w:broken'", "/w:body[1]", "/word/document.xml");
    const pipeline = makePipeline({
      staging: { passed: false, message: "", findings: [quirkA, fresh] },
      live: { passed: false, message: "", findings: [quirkA] }
    });
    const report = await verify(pipeline, "live");
    expect(report.structural.status).toBe("fail");
    expect(report.structural.issues.some((i) => i.code === "schema-invalid")).toBe(true);
    expect(report.structural.issues.some((i) => i.message.includes("w:broken"))).toBe(true);
  });

  it("a clean candidate still passes outright", async () => {
    const pipeline = makePipeline({
      staging: { passed: true, message: "Validation passed: no errors found.", findings: [] }
    });
    const report = await verify(pipeline, "live");
    expect(report.structural.status).toBe("pass");
    expect(report.confidence).toBe("engine");
  });

  it("an unresolvable base revision fails closed (every finding treated as new)", async () => {
    const quirkA = finding("[Schema] uiPriority", "/w:styles[1]/w:style[2]", "/word/styles.xml");
    const pipeline = makePipeline({
      staging: { passed: false, message: "", findings: [quirkA] },
      live: { passed: false, message: "", findings: [quirkA] }
    });
    // baseRevision.artifactRef "gone" is unresolvable → resolvePath throws.
    const report = await verify(pipeline, "gone");
    expect(report.structural.status).toBe("fail");
  });
});
