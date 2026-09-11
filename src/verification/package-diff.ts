/**
 * PackageDiff (§80): raw part diff → safe normalization → relationship diff →
 * risk classification. Classification feeds the L4 mutation-aware check.
 */

import type {
  PackageDiffEntry,
  PackageDiffResult,
  PartDiffClass
} from "../contracts/verification.js";
import type { PackageManifest } from "../contracts/artifact.js";
import { canonicalizeXml } from "../support/xml-lite.js";

/** Parts whose unexpected modification is structurally high-risk. */
const HIGH_RISK_PARTS = [
  /[\/_]theme\d*\.xml$/,
  /[\/_]slideMaster\d*\.xml$/,
  /[\/_]slideLayout\d*\.xml$/,
  /[\/_]sheetM?asters?\d*\.xml$/, // xlsx workbook protection lives near masters/styles
  /styles\.xml$/,
  /settings\.xml$/,
  /[\/_]numbering\.xml$/,
  /\[Content_Types\]\.xml$/
];

const RELATIONSHIP_PART = /\.rels$/;

/** Expected side-effect parts an agent editing slide/sheet N may touch. */
function expectedSideParts(changedTargets: string[]): RegExp[] {
  // Engine saves stamp document properties — always expected side effects.
  const patterns: RegExp[] = [/docProps\//];
  for (const target of changedTargets) {
    const slideMatch = target.match(/\/slide\[(\d+)\]/);
    if (slideMatch) {
      patterns.push(new RegExp(`ppt/slides/slide${slideMatch[1]}\\.xml$`));
      patterns.push(new RegExp(`ppt/slides/_rels/slide${slideMatch[1]}\\.xml\\.rels$`));
      continue;
    }
    const sheetMatch = target.match(/\/sheet(?:\[(\d+)\])?/);
    if (sheetMatch) {
      // Sheet edits may touch sharedStrings, worksheets, calcChain.
      patterns.push(/xl\/worksheets\//, /xl\/sharedStrings\.xml$/, /xl\/calcChain\.xml$/, /xl\/workbook\.xml$/);
      continue;
    }
    if (target.includes("paragraph") || target.includes("block") || target.startsWith("/body")) {
      patterns.push(/word\/document\.xml$/, /word\/_rels\/document\.xml\.rels$/);
    }
  }
  return patterns;
}

export interface PackageDiffInput {
  baseManifest: PackageManifest;
  candidateManifest: PackageManifest;
  /** Raw bytes accessor for normalization comparison (bounded by size cap). */
  readPart?: (manifest: PackageManifest, name: string, which: "base" | "candidate") => Promise<Buffer | undefined>;
  /** Mutation scope targets (e.g. ["/slide[3]"]) for EXPECTED classification. */
  changedTargets?: string[];
}

const NORMALIZATION_COMPARE_CAP = 8 * 1024 * 1024;

export async function diffPackages(input: PackageDiffInput): Promise<PackageDiffResult> {
  const baseEntries = new Map(input.baseManifest.entries.map((e) => [e.name, e]));
  const candidateEntries = new Map(input.candidateManifest.entries.map((e) => [e.name, e]));
  const allNames = new Set([...baseEntries.keys(), ...candidateEntries.keys()]);

  const expectedPatterns = expectedSideParts(input.changedTargets ?? []);
  const entries: PackageDiffEntry[] = [];
  const relationshipChanges: string[] = [];
  const highRiskParts: string[] = [];
  const summary: Record<PartDiffClass, number> = {
    UNCHANGED: 0,
    NORMALIZATION_ONLY: 0,
    EXPECTED: 0,
    UNEXPECTED_LOW_RISK: 0,
    UNEXPECTED_HIGH_RISK: 0
  };

  for (const name of allNames) {
    const base = baseEntries.get(name);
    const candidate = candidateEntries.get(name);

    if (base && candidate && base.crc32 === candidate.crc32 && base.size === candidate.size) {
      summary.UNCHANGED++;
      continue;
    }

    let change: PackageDiffEntry["change"];
    if (base && candidate) change = "modified";
    else if (candidate) change = "added";
    else change = "removed";

    let classification: PartDiffClass;
    if (expectedPatterns.some((re) => re.test(name))) {
      classification = "EXPECTED";
    } else if (HIGH_RISK_PARTS.some((re) => re.test(name))) {
      classification = "UNEXPECTED_HIGH_RISK";
    } else if (change === "modified" && input.readPart && candidate && base) {
      classification = await classifyModified(name, base.size, candidate.size, input, input.readPart);
    } else {
      classification = "UNEXPECTED_LOW_RISK";
    }

    if (RELATIONSHIP_PART.test(name)) {
      relationshipChanges.push(name);
      if (classification === "EXPECTED") {
        // Relationship edits tied to expected content changes stay expected.
      } else {
        classification = "UNEXPECTED_HIGH_RISK";
      }
    }
    if (classification === "UNEXPECTED_HIGH_RISK") highRiskParts.push(name);

    entries.push({ name, change, classification });
    summary[classification]++;
  }

  return { entries, summary, relationshipChanges, highRiskParts };
}

async function classifyModified(
  name: string,
  baseSize: number,
  candidateSize: number,
  input: PackageDiffInput,
  readPart: NonNullable<PackageDiffInput["readPart"]>
): Promise<PartDiffClass> {
  if (!name.endsWith(".xml") && !name.endsWith(".rels")) return "UNEXPECTED_LOW_RISK";
  if (baseSize > NORMALIZATION_COMPARE_CAP || candidateSize > NORMALIZATION_COMPARE_CAP) {
    return "UNEXPECTED_LOW_RISK";
  }
  const [baseBytes, candidateBytes] = await Promise.all([
    readPart(input.baseManifest, name, "base"),
    readPart(input.candidateManifest, name, "candidate")
  ]);
  if (!baseBytes || !candidateBytes) return "UNEXPECTED_LOW_RISK";
  const baseCanon = canonicalizeXml(baseBytes.toString("utf8"));
  const candidateCanon = canonicalizeXml(candidateBytes.toString("utf8"));
  if (baseCanon === candidateCanon) return "NORMALIZATION_ONLY";
  if (HIGH_RISK_PARTS.some((re) => re.test(name))) return "UNEXPECTED_HIGH_RISK";
  return "UNEXPECTED_LOW_RISK";
}
