/**
 * PortableReviewView (§69, §85): review surface for generic MCP hosts —
 * before/after outline, verification status, accept/reject actions. It is a
 * view + action bundle, never the main editor.
 */

import type { OfficeRuntimeService } from "../../runtime/service/office-runtime-service.js";
import type { PreviewOutline } from "../../contracts/preview.js";

export interface ReviewDiff {
  base: { revisionId: string; contentHash: string };
  candidate: { candidateId: string; state: string; contentHash?: string };
  /** Outline before (committed) and after (candidate) for side-by-side view. */
  before: PreviewOutline;
  after: PreviewOutline;
  verification?: { confidence: string; checks: Array<{ layer: string; status: string }> };
}

export class PortableReviewView {
  constructor(private readonly service: OfficeRuntimeService) {}

  async build(sessionId: string): Promise<ReviewDiff> {
    const session = this.service.getSession(sessionId);
    if (!session) throw new Error(`unknown session ${sessionId}`);
    const candidate = session.candidate;
    if (!candidate) throw new Error(`no active candidate on session ${sessionId}`);

    const basePreview = await this.service.preview({ artifactRef: session.artifactRef, priority: "visible" });
    const candidatePreview = await this.service.preview({
      artifactRef: candidate.artifactRef,
      priority: "visible"
    });
    return {
      base: { revisionId: session.committedRevision.revisionId, contentHash: session.committedRevision.contentHash },
      candidate: { candidateId: candidate.candidateId, state: candidate.state, contentHash: candidate.currentHash },
      before: basePreview.model.outline,
      after: candidatePreview.model.outline,
      verification: candidate.verification
        ? {
            confidence: candidate.verification.confidence,
            checks: [
              { layer: candidate.verification.structural.layer, status: candidate.verification.structural.status },
              { layer: candidate.verification.package.layer, status: candidate.verification.package.status },
              { layer: candidate.verification.semantic.layer, status: candidate.verification.semantic.status },
              { layer: candidate.verification.visual.layer, status: candidate.verification.visual.status }
            ]
          }
        : undefined
    };
  }

  async accept(sessionId: string): Promise<{ revisionId: string }> {
    const session = this.service.getSession(sessionId);
    if (!session?.candidate) throw new Error("nothing to accept");
    const result = await this.service.acceptCandidate(sessionId, session.candidate.candidateId);
    return { revisionId: result.revisionId };
  }

  async reject(sessionId: string): Promise<{ rejected: true }> {
    const session = this.service.getSession(sessionId);
    if (!session?.candidate) throw new Error("nothing to reject");
    await this.service.rejectCandidate(sessionId, session.candidate.candidateId);
    return { rejected: true };
  }
}
