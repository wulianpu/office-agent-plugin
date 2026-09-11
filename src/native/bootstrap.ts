/**
 * Native bootstrap (§7–§8): the OfficeRuntimeService lives in an Electron
 * utilityProcess in the full product. This module is the embeddable entry:
 * a MessagePort-style RPC handler plus a plain-Node launcher, so the same
 * service runs inside Electron's utilityProcess.fork() or standalone.
 */

import type { OfficeRuntimeService } from "../runtime/service/office-runtime-service.js";
import { OfficeMcpTools } from "../mcp/tools.js";

export type RuntimeRequest =
  | { kind: "open"; path: string }
  | { kind: "preview"; artifactRef: string; priority?: "visible" | "prefetch" | "background" }
  | { kind: "session.open"; artifactRef: string }
  | { kind: "session.close"; sessionId: string }
  | { kind: "agent.begin"; sessionId: string; intent: string; destructiveAllowed?: boolean; allowedTargets?: string[] }
  | { kind: "agent.edit"; taskId: string; idempotencyKey: string; items: unknown[] }
  | { kind: "agent.flush"; taskId: string }
  | { kind: "agent.finalize"; taskId: string }
  | { kind: "candidate.accept"; sessionId: string }
  | { kind: "candidate.reject"; sessionId: string }
  | { kind: "capabilities" };

export type RuntimeResponse =
  | { kind: "ok"; data: unknown }
  | { kind: "error"; code: string; message: string };

/**
 * Handle one RPC request against a live service. MessagePort transports call
 * this directly; the CLI/demo uses it too, keeping one dispatch table.
 */
export async function handleRuntimeRequest(
  service: OfficeRuntimeService,
  request: RuntimeRequest
): Promise<RuntimeResponse> {
  try {
    switch (request.kind) {
      case "open": {
        const ref = await service.registerArtifact(request.path);
        return { kind: "ok", data: { artifactRef: ref } };
      }
      case "preview": {
        const result = await service.preview({
          artifactRef: request.artifactRef,
          priority: request.priority ?? "visible"
        });
        return { kind: "ok", data: result };
      }
      case "session.open": {
        const session = await service.openSession(request.artifactRef);
        return {
          kind: "ok",
          data: {
            sessionId: session.sessionId,
            documentId: session.documentId,
            revisionId: session.committedRevision.revisionId
          }
        };
      }
      case "session.close": {
        await service.closeSession(request.sessionId);
        return { kind: "ok", data: { closed: true } };
      }
      case "capabilities":
        return { kind: "ok", data: new OfficeMcpTools(service).capabilities() };
      default:
        return { kind: "error", code: "unsupported", message: `request ${request.kind} requires the high-level plugin API` };
    }
  } catch (error) {
    return {
      kind: "error",
      code: (error as { code?: string }).code ?? "runtime-error",
      message: String((error as Error)?.message ?? error)
    };
  }
}
