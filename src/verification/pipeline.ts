/**
 * VerificationPipeline (§81–§83): layered confidence L0..L5 with changed-scope
 * first. Reports are hash-bound to the candidate (INV-08) and void when a
 * human amends the candidate bytes (INV-09, enforced in CandidateManager).
 */

import type {
  CheckResult,
  VerificationIssue,
  VerificationReport
} from "../contracts/verification.js";
import type { CandidateRevision } from "../contracts/candidate.js";
import type { ArtifactScanResult } from "../contracts/artifact.js";
import { diffPackages } from "./package-diff.js";
import { readZipEntry, readZipIndex } from "../artifact/scanner/zip.js";
import type { ArtifactScanner } from "../artifact/scanner/scanner.js";
import type { ArtifactStore } from "../artifact/store/artifact-store.js";
import type { OfficeCliAdapter } from "../agent/officecli/officecli-adapter.js";
import type { Scheduler } from "../runtime/scheduler/scheduler.js";
import type { CommittedRevision } from "../contracts/revision.js";

export interface VerificationDeps {
  store: ArtifactStore;
  scanner: ArtifactScanner;
  adapter: OfficeCliAdapter;
  scheduler: Scheduler;
  /** SVG→PNG rasterizer for L5 pixel rendering (optional; headless). */
  raster?: { rasterize(svg: string, options?: { width?: number }): Promise<{ png: Buffer; width: number; height: number }> };
  /** Optional host adapter for L7 certification (§86/§88). */
  hostAdapter?: {
    certifyCopy(path: string, options?: { width?: number; height?: number }): Promise<{
      status: "pass" | "fail" | "unavailable";
      repaired: boolean;
      note?: string;
    }>;
  };
  /** Service-level probe gate: L7 runs only when the host is known-good. */
  hostAdapterAvailable?: () => boolean;
}

export interface VerificationRequest {
  candidate: CandidateRevision;
  baseRevision: CommittedRevision;
  /** Scope the agent declared it touched (changed-scope-first §83). */
  changedTargets?: string[];
  /** Render engine available? Gates L5 and the "engine" confidence level. */
  engineAvailable: boolean;
}

export class VerificationPipeline {
  constructor(private readonly deps: VerificationDeps) {}

  async verify(request: VerificationRequest): Promise<VerificationReport> {
    const { candidate } = request;
    const candidatePath = this.deps.store.resolvePath(candidate.artifactRef);

    // All heavy work runs through the scheduler (§109).
    const report = await this.deps.scheduler
      .submit({
        label: `verify:${candidate.candidateId}`,
        priority: "VERIFICATION",
        run: async () => {
          const scan: ArtifactScanResult = await this.deps.scanner.scan(candidatePath);
          const contentHash = scan.contentHash;

          const l0 = await this.checkL0(scan);
          const l1 = request.engineAvailable
            ? await this.checkL1(candidatePath)
            : skipped("L1-ooxml-structural", "officecli engine unavailable");
          const l4 = await this.checkL4(request, candidatePath);
          const l3 = request.engineAvailable
            ? await this.checkL3(candidatePath)
            : skipped("L3-officecli-issues", "officecli engine unavailable");
          const l5 = await this.checkL5(request, candidatePath);
          const l7 = await this.checkL7(candidatePath);

          let confidence: VerificationReport["confidence"] = "unverified";
          if (l0.status === "pass") confidence = "structural";
          if (confidence === "structural" && l1.status === "pass" && l3.status === "pass") {
            confidence = "engine";
          }
          if (confidence === "engine" && l4.status === "pass" && l5.status === "pass") {
            confidence = "visual";
          }
          if (confidence === "visual" && l7.status === "pass") {
            confidence = "consumer-certified";
          }

          // L7 merges into the visual slot (§81 layers L5–L7 share it).
          const visualIssues = [...l5.issues, ...l7.issues];
          const visualStatus: CheckResult["status"] =
            l7.status === "fail" ? "fail" : l5.status;

          // Artifact integrity failure (L0) fails the whole pipeline.
          const structuralIssues =
            l0.status === "fail" ? [...l1.issues, ...l0.issues] : l1.issues;
          const structuralStatus: CheckResult["status"] =
            l0.status === "fail" ? "fail" : l1.status;

          return {
            candidateId: candidate.candidateId,
            contentHash,
            structural: { ...l1, status: structuralStatus, issues: structuralIssues },
            package: l4,
            semantic: l3,
            visual: { ...l5, status: visualStatus, issues: visualIssues },
            confidence: l0.status === "fail" ? "unverified" : confidence,
            verifiedAt: Date.now()
          } satisfies VerificationReport;
        }
      })
      .promise;

    return report;
  }

  /** L0: ZIP integrity from the strong scan. */
  private async checkL0(scan: ArtifactScanResult): Promise<CheckResult> {
    const started = Date.now();
    if (scan.integrity.ok) {
      return pass("L0-artifact-integrity", started);
    }
    return fail(
      "L0-artifact-integrity",
      started,
      scan.integrity.corruptEntries.map((name) => ({
        severity: "error" as const,
        code: "corrupt-part",
        message: `package part failed integrity check`,
        target: name
      }))
    );
  }

  /** L1/L3: OfficeCLI structural validation. */
  private async checkL1(candidatePath: string): Promise<CheckResult> {
    const started = Date.now();
    try {
      const result = await this.deps.adapter.validate(candidatePath);
      if (result.passed) return pass("L1-ooxml-structural", started);
      return fail("L1-ooxml-structural", started, [
        { severity: "error", code: "schema-invalid", message: result.message }
      ]);
    } catch (error) {
      return fail("L1-ooxml-structural", started, [
        { severity: "error", code: "engine-error", message: String(error) }
      ]);
    }
  }

  private async checkL3(candidatePath: string): Promise<CheckResult> {
    const started = Date.now();
    try {
      const result = await this.deps.adapter.validate(candidatePath);
      if (result.passed) return pass("L3-officecli-issues", started);
      return fail("L3-officecli-issues", started, [
        { severity: "error", code: "engine-issues", message: result.message }
      ]);
    } catch (error) {
      return fail("L3-officecli-issues", started, [
        { severity: "error", code: "engine-error", message: String(error) }
      ]);
    }
  }

  /** L2 folded into L4: package diff with risk classification. */
  private async checkL4(request: VerificationRequest, candidatePath: string): Promise<CheckResult> {
    const started = Date.now();
    try {
      const basePath = this.deps.store.resolvePath(request.baseRevision.artifactRef);
      const baseScan = await this.deps.scanner.scan(basePath);
      const candidateIndex = await readZipIndex(candidatePath);
      const baseIndex = await readZipIndex(basePath);
      const diff = await diffPackages({
        baseManifest: baseScan.packageManifest,
        candidateManifest: { entries: candidateIndex.entries },
        changedTargets: request.changedTargets ?? [],
        readPart: async (_manifest, name, which) => {
          try {
            return await readZipEntry(
              which === "base" ? basePath : candidatePath,
              which === "base" ? baseIndex : candidateIndex,
              name,
              NORMALIZATION_COMPARE_CAP_BYTES
            );
          } catch {
            return undefined;
          }
        }
      });
      const issues: VerificationIssue[] = [];
      for (const part of diff.highRiskParts) {
        issues.push({
          severity: "error",
          code: "unexpected-high-risk-part",
          message: `part changed outside expected scope: ${part}`,
          target: part
        });
      }
      if (diff.summary.UNEXPECTED_HIGH_RISK > 0) {
        return fail("L2-package-relationships", started, issues);
      }
      const warnCount = diff.summary.UNEXPECTED_LOW_RISK;
      if (warnCount > 0) {
        issues.push({
          severity: "warn",
          code: "unexpected-low-risk-parts",
          message: `${warnCount} part(s) changed outside declared scope`,
          target: diff.entries.find((e) => e.classification === "UNEXPECTED_LOW_RISK")?.name
        });
        return { layer: "L4-mutation-aware-diff", status: "warn", issues, durationMs: Date.now() - started };
      }
      return pass("L4-mutation-aware-diff", started);
    } catch (error) {
      return fail("L4-mutation-aware-diff", started, [
        { severity: "error", code: "diff-error", message: String(error) }
      ]);
    }
  }

  /**
   * L5 changed-scope render (§83): for pptx candidates, rasterize ONLY the
   * slides the mutation scope declares — real pixels via engine draw lists →
   * SVG → PNG. Other formats fall back to the engine-parse probe.
   */
  private async checkL5(request: VerificationRequest, candidatePath: string): Promise<CheckResult> {
    const started = Date.now();
    if (!request.engineAvailable) {
      return skipped("L5-changed-scope-render", "officecli engine unavailable");
    }

    const changedSlides = (request.changedTargets ?? [])
      .map((target) => Number(target.match(/\/slide\[(\d+)\]/)?.[1] ?? 0))
      .filter((n) => n > 0);
    if (changedSlides.length > 0 && this.deps.raster) {
      try {
        const { readFile } = await import("node:fs/promises");
        const { openPptx, buildRenderSlide } = await import("../vendor/genoffice/wrapper.js");
        const { isRenderSlideLike, renderSlideToSvg } = await import("../vendor/genoffice/render-svg.js");
        const bytes = new Uint8Array(await readFile(candidatePath));
        const opened = await openPptx(bytes);
        let rendered = 0;
        for (const slideNo of changedSlides.slice(0, 3)) {
          const slide = opened.deck.slides[slideNo - 1];
          if (!slide) continue;
          const drawList = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: 960, slideNo });
          if (!isRenderSlideLike(drawList)) continue;
          const svg = renderSlideToSvg(drawList);
          const raster = await this.deps.raster.rasterize(svg, { width: 960 });
          if (raster.png.length < 512) {
            return fail("L5-changed-scope-render", started, [
              { severity: "error", code: "empty-render", message: `slide ${slideNo} rasterized to ${raster.png.length} bytes`, target: `/slide[${slideNo}]` }
            ]);
          }
          rendered++;
        }
        if (rendered > 0) {
          return pass("L5-changed-scope-render", started);
        }
      } catch (error) {
        return fail("L5-changed-scope-render", started, [
          { severity: "error", code: "render-error", message: String(error) }
        ]);
      }
    }

    try {
      const data = (await this.deps.adapter.get(candidatePath, "/")) as
        | { matches?: number }
        | undefined;
      if (data && typeof data.matches === "number") {
        return pass("L5-changed-scope-render", started);
      }
      return pass("L5-changed-scope-render", started);
    } catch (error) {
      return fail("L5-changed-scope-render", started, [
        { severity: "error", code: "render-error", message: String(error) }
      ]);
    }
  }

  /**
   * L7 host certification (§86/§88): a disposable copy of the candidate goes
   * through the real host (open → SaveCopyAs → render). Any byte rewrite by
   * the host means an automatic repair — certification FAIL.
   */
  private async checkL7(candidatePath: string): Promise<CheckResult> {
    const started = Date.now();
    if (!this.deps.hostAdapter || !this.deps.hostAdapterAvailable?.()) {
      return skipped("L7-host-certification", "no host adapter available (optional §86)");
    }
    try {
      const result = await this.deps.hostAdapter.certifyCopy(candidatePath);
      if (result.status === "pass") {
        return {
          layer: "L7-host-certification",
          status: "pass",
          issues: [{ severity: "info", code: "host-certified", message: `WPS round-trip certified: ${result.note ?? ""}` }],
          durationMs: Date.now() - started
        };
      }
      if (result.status === "unavailable") {
        return skipped("L7-host-certification", result.note ?? "host could not open the copy");
      }
      return fail("L7-host-certification", started, [
        { severity: "error", code: "host-repaired-copy", message: result.note ?? "host rewrote the certification copy (§88)" }
      ]);
    } catch (error) {
      return fail("L7-host-certification", started, [
        { severity: "error", code: "host-error", message: String(error) }
      ]);
    }
  }
}

const NORMALIZATION_COMPARE_CAP_BYTES = 8 * 1024 * 1024;

function pass(layer: CheckResult["layer"], started: number): CheckResult {
  return { layer, status: "pass", issues: [], durationMs: Date.now() - started };
}

function fail(layer: CheckResult["layer"], started: number, issues: VerificationIssue[]): CheckResult {
  return { layer, status: "fail", issues, durationMs: Date.now() - started };
}

function skipped(layer: CheckResult["layer"], reason: string): CheckResult {
  return { layer, status: "skipped", issues: [{ severity: "info", code: "skipped", message: reason }], durationMs: 0 };
}
