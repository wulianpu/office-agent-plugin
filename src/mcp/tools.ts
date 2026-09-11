/**
 * MCP tool core (§57–§62): transport-agnostic implementation of the six
 * office.* tools. Agents receive document capabilities, never paths (INV-12) —
 * every result string passes through path sanitization.
 */

import type {
  OfficeCapabilitiesResult,
  OfficeCapabilitiesResult as CapabilitiesResult,
  OfficeEditInput,
  OfficeEditResult,
  OfficeInspectInput,
  OfficeInspectResult,
  OfficeQueryInput,
  OfficeQueryResult,
  OfficeRenderInput,
  OfficeRenderResult,
  OfficeVerifyInput,
  OfficeVerifyResult
} from "../contracts/mcp.js";
import type { CapabilityStatus } from "../contracts/mcp.js";
import { OfficeRuntimeError } from "../contracts/document.js";
import type { OfficeRuntimeService } from "../runtime/service/office-runtime-service.js";
import { newCommandId } from "../support/ids.js";
import type { PreviewOutline } from "../contracts/preview.js";

export class OfficeMcpTools {
  constructor(private readonly service: OfficeRuntimeService) {}

  /** Sanitize output: physical paths never leave the runtime (INV-12). */
  private sanitize<T>(value: T): T {
    const replacer = (text: string): string => {
      let out = text;
      for (const artifact of this.service.repos.loadArtifacts()) {
        out = out.split(artifact.path).join(`<artifact:${artifact.ref}>`);
      }
      return out;
    };
    return JSON.parse(
      JSON.stringify(value, (_key, val) => (typeof val === "string" ? replacer(val) : val))
    ) as T;
  }

  private requireSession(sessionId: string) {
    const session = this.service.getSession(sessionId);
    if (!session) throw new OfficeRuntimeError("session-closed", `no open session: ${sessionId}`);
    return session;
  }

  async inspect(input: OfficeInspectInput): Promise<OfficeInspectResult> {
    const session = this.requireSession(input.sessionId);
    if (!this.service.isEngineAvailable()) {
      throw new OfficeRuntimeError("recovery-required", "officecli engine unavailable");
    }
    const path = input.path ?? "/";
    // Read from the candidate when one is active (agent sees its own work).
    const targetRef =
      session.candidate?.artifactRef && session.candidate.state !== "failed"
        ? session.candidate.artifactRef
        : session.artifactRef;
    const data = await this.service.officecli.get(this.service.store.resolvePath(targetRef), path);
    return this.sanitize({ format: session.format, node: data, path });
  }

  async query(input: OfficeQueryInput): Promise<OfficeQueryResult> {
    const session = this.requireSession(input.sessionId);
    if (!this.service.isEngineAvailable()) {
      throw new OfficeRuntimeError("recovery-required", "officecli engine unavailable");
    }
    const targetRef =
      session.candidate?.artifactRef && session.candidate.state !== "failed"
        ? session.candidate.artifactRef
        : session.artifactRef;
    const data = (await this.service.officecli.query(
      this.service.store.resolvePath(targetRef),
      input.selector
    )) as { results?: Array<{ path: string; preview?: string; type?: string }>; matches?: number } | undefined;
    const matches = (data?.results ?? []).map((r) => ({
      path: r.path,
      summary: { type: r.type ?? "unknown", preview: r.preview ?? "" }
    }));
    return this.sanitize({ matches, truncated: (data?.matches ?? 0) > matches.length });
  }

  async edit(input: OfficeEditInput): Promise<OfficeEditResult> {
    const session = this.requireSession(input.sessionId);

    const mutationScope = {
      intent: input.intent,
      destructiveAllowed: input.scope?.destructiveAllowed ?? false,
      allowedTargets: input.scope?.allowedTargets,
      allowedParts: input.scope?.allowedParts
    };

    // Reuse the live agent task for the active candidate, or open one.
    let taskContext = session.candidate && session.candidate.state !== "failed"
      ? this.service.agent.taskForCandidate(session.candidate.candidateId)
      : undefined;
    if (!taskContext) {
      taskContext = await this.service.beginAgentTask(session.sessionId, mutationScope);
    }

    const receipt = await this.service.executeAgentMutation(
      taskContext,
      {
        commandId: newCommandId(),
        idempotencyKey: input.idempotencyKey,
        approved: false,
        payload: input.items
      }
    );
    return this.sanitize({
      taskContext: {
        taskId: taskContext.taskId,
        candidateId: taskContext.candidateId,
        baseRevisionId: taskContext.baseRevisionId,
        fencingToken: taskContext.fencingToken.toString()
      },
      receipts: [receipt],
      notes: []
    });
  }

  async render(input: OfficeRenderInput): Promise<OfficeRenderResult> {
    const session = this.requireSession(input.sessionId);
    const targetRef =
      session.candidate?.artifactRef && session.candidate.state !== "failed"
        ? session.candidate.artifactRef
        : session.artifactRef;
    const preview = await this.service.preview({
      artifactRef: targetRef,
      priority: "visible"
    });
    const outline: PreviewOutline = preview.model.outline;
    const sections: Array<{ label: string; text: string }> = [];
    if (outline.kind === "docx") {
      outline.blocks
        .filter((b) => b.text.trim().length > 0)
        .slice(0, 200)
        .forEach((b) => sections.push({ label: `block ${b.index}${b.style ? ` (${b.style})` : ""}`, text: b.text }));
    } else if (outline.kind === "xlsx") {
      for (const sheet of outline.sheets) {
        sections.push({
          label: `sheet ${sheet.name}`,
          text: sheet.window.map((row) => row.join(" | ")).join("\n")
        });
      }
    } else {
      for (const slide of outline.slides) {
        sections.push({
          label: `slide ${slide.index}`,
          text: slide.shapes.map((s) => s.text ?? s.name ?? "").filter(Boolean).join(" / ")
        });
      }
    }
    return this.sanitize({ mode: input.mode ?? "outline", sections });
  }

  async verify(input: OfficeVerifyInput): Promise<OfficeVerifyResult> {
    const session = this.requireSession(input.sessionId);
    const candidate = input.candidateId
      ? this.service.candidates.require(input.candidateId)
      : session.candidate;
    if (!candidate || !candidate.currentHash) {
      throw new OfficeRuntimeError("candidate-not-ready", "candidate not flushed");
    }
    if (candidate.verification && candidate.verification.contentHash === candidate.currentHash) {
      return this.sanitize({ report: candidate.verification });
    }
    throw new OfficeRuntimeError("verification-stale", "candidate has no bound verification; run the pipeline");
  }

  capabilities(): OfficeCapabilitiesResult {
    const engine = this.service.isEngineAvailable();
    const editorStatus = (format: "docx" | "xlsx" | "pptx"): CapabilityStatus => {
      if (format === "xlsx") {
        if (this.service.isXlsxSidecarAvailable()) {
          return {
            status: "available",
            engine: "genoffice-sheets",
            reason: "Rust sidecar (calamine + IronCalc) serves bounded viewports; visual Univer canvas awaits an Electron host"
          };
        }
        return {
          status: "degraded",
          engine: "basic",
          reason: "xlsx Rust sidecar not built (npm run sidecar:build); streaming viewport path active"
        };
      }
      if (this.service.isGenOfficeAvailable(format)) {
        return {
          status: "available",
          engine: "genoffice",
          reason: "engine-accurate parse/preview via vendored GenOffice engine; visual canvas awaits an Electron host (§P7)"
        };
      }
      return { status: "degraded", engine: "basic", reason: "GenOffice vendor bundle absent; run npm run build:vendor" };
    };
    const agentStatus: CapabilityStatus = engine
      ? { status: "available", engine: "officecli" }
      : { status: "unavailable", engine: "officecli", reason: "officecli not resolvable on PATH" };
    const verifyStatus: CapabilityStatus = engine
      ? { status: "available", engine: "officecli" }
      : { status: "degraded", engine: "internal", reason: "structural-only verification without engine" };
    return {
      capabilities: (["docx", "xlsx", "pptx"] as const).map((format) => ({
        format,
        editor: editorStatus(format),
        agent: agentStatus,
        verification: verifyStatus
      })),
      hostAdapters: [
        {
          id: "powerpoint",
          status: this.service.isHostAvailable("powerpoint")
            ? { status: "available", engine: "PowerPoint-COM" }
            : { status: "unavailable", engine: "none", reason: "PowerPoint COM registered but the application server fails to start (stale registration)" }
        },
        {
          id: "wps",
          status: this.service.isHostAvailable("wps")
            ? { status: "available", engine: "KWPP-COM" }
            : { status: "unavailable", engine: "none", reason: "WPS Office not installed or KWPP COM unavailable" }
        }
      ],
      offline: true
    };
  }
}

export type { CapabilitiesResult };
