import { describe, expect, it } from "vitest";
import { diffPackages } from "../../src/verification/package-diff.js";
import type { PackageManifest, PackageManifestEntry } from "../../src/contracts/artifact.js";

const entry = (name: string, size = 100, crc = 1): PackageManifestEntry => ({
  name,
  compressedSize: size,
  size,
  crc32: crc,
  method: 8
});

const manifest = (entries: PackageManifestEntry[]): PackageManifest => ({ entries });

describe("PackageDiff (§80)", () => {
  it("classifies untouched parts UNCHANGED and scoped parts EXPECTED", async () => {
    const diff = await diffPackages({
      baseManifest: manifest([entry("ppt/slides/slide1.xml"), entry("ppt/slides/slide2.xml")]),
      candidateManifest: manifest([
        entry("ppt/slides/slide1.xml"),
        entry("ppt/slides/slide2.xml", 200, 2) // slide2 changed
      ]),
      changedTargets: ["/slide[2]"]
    });
    expect(diff.summary.UNCHANGED).toBe(1);
    const changed = diff.entries.find((e) => e.change === "modified");
    expect(changed?.classification).toBe("EXPECTED");
  });

  it("flags unexpected theme/master edits HIGH RISK", async () => {
    const diff = await diffPackages({
      baseManifest: manifest([entry("ppt/theme/theme1.xml", 50, 1), entry("ppt/slides/slide1.xml")]),
      candidateManifest: manifest([entry("ppt/theme/theme1.xml", 50, 2), entry("ppt/slides/slide1.xml")]),
      changedTargets: ["/slide[1]"]
    });
    expect(diff.highRiskParts).toContain("ppt/theme/theme1.xml");
    expect(diff.summary.UNEXPECTED_HIGH_RISK).toBe(1);
  });

  it("detects NORMALIZATION_ONLY via canonicalization", async () => {
    const baseXml = `<a x="1" y="2"><b>t</b>  </a>`;
    const candidateXml = `<a y='2' x="1"><b>t</b></a>`;
    const diff = await diffPackages({
      baseManifest: manifest([entry("word/document.xml", 50, 1)]),
      candidateManifest: manifest([entry("word/document.xml", 50, 2)]),
      readPart: async (_m, name, which) =>
        name === "word/document.xml" ? Buffer.from(which === "base" ? baseXml : candidateXml) : undefined
    });
    const entryResult = diff.entries.find((e) => e.name === "word/document.xml");
    expect(entryResult?.classification).toBe("NORMALIZATION_ONLY");
  });

  it("treats surprise .rels edits as high risk", async () => {
    const diff = await diffPackages({
      baseManifest: manifest([entry("ppt/slides/_rels/slide1.xml.rels", 30, 1)]),
      candidateManifest: manifest([entry("ppt/slides/_rels/slide1.xml.rels", 30, 2)]),
      changedTargets: ["/slide[9]"] // unrelated scope
    });
    expect(diff.highRiskParts).toContain("ppt/slides/_rels/slide1.xml.rels");
  });
});
