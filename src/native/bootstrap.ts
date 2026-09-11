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
  request: RuntimeRequest,
  plugin?: {
    beginAgentTask(sessionId: string, scope: { intent: string; destructiveAllowed?: boolean; allowedTargets?: string[] }): Promise<{ taskId: string; candidateId: string; baseRevisionId: string; fencingToken: bigint }>;
    executeAgentMutation(task: unknown, command: { commandId: string; idempotencyKey: string; payload: unknown[] }): Promise<unknown>;
    flushAgentCandidate(task: unknown): Promise<string>;
    finalizeAgentTask(task: unknown): Promise<void>;
    acceptCandidate(sessionId: string): Promise<{ revisionId: string; contextPromoted: boolean }>;
    rejectCandidate(sessionId: string): Promise<void>;
  }
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
      case "agent.begin": {
        if (!plugin) return { kind: "error", code: "unsupported", message: "agent RPC requires the plugin facade" };
        const task = await plugin.beginAgentTask(request.sessionId, {
          intent: request.intent ?? "rpc",
          destructiveAllowed: request.destructiveAllowed ?? false,
          allowedTargets: request.allowedTargets
        });
        return { kind: "ok", data: { taskId: task.taskId, candidateId: task.candidateId, baseRevisionId: task.baseRevisionId, fencingToken: task.fencingToken.toString() } };
      }
      case "agent.edit": {
        if (!plugin) return { kind: "error", code: "unsupported", message: "agent RPC requires the plugin facade" };
        const tasks = service.agent.activeTasks();
        const task = tasks.find((t) => t.taskId === request.taskId);
        if (!task) return { kind: "error", code: "not_found", message: `unknown task ${request.taskId}` };
        const receipt = await plugin.executeAgentMutation(task, {
          commandId: `rpc-${request.taskId}-${request.idempotencyKey}`,
          idempotencyKey: request.idempotencyKey,
          payload: request.items
        });
        return { kind: "ok", data: receipt };
      }
      case "agent.flush":
      case "agent.finalize": {
        if (!plugin) return { kind: "error", code: "unsupported", message: "agent RPC requires the plugin facade" };
        const task = service.agent.activeTasks().find((t) => t.taskId === request.taskId);
        if (!task) return { kind: "error", code: "not_found", message: `unknown task ${request.taskId}` };
        if (request.kind === "agent.flush") {
          return { kind: "ok", data: { contentHash: await plugin.flushAgentCandidate(task) } };
        }
        await plugin.finalizeAgentTask(task);
        return { kind: "ok", data: { finalized: true } };
      }
      case "candidate.accept": {
        if (!plugin) return { kind: "error", code: "unsupported", message: "candidate RPC requires the plugin facade" };
        const accepted = await plugin.acceptCandidate(request.sessionId);
        return { kind: "ok", data: accepted };
      }
      case "candidate.reject": {
        if (!plugin) return { kind: "error", code: "unsupported", message: "candidate RPC requires the plugin facade" };
        await plugin.rejectCandidate(request.sessionId);
        return { kind: "ok", data: { rejected: true } };
      }
      default:
        return { kind: "error", code: "unsupported", message: `unsupported request ${(request as { kind: string }).kind}` };
    }
  } catch (error) {
    return {
      kind: "error",
      code: (error as { code?: string }).code ?? "runtime-error",
      message: String((error as Error)?.message ?? error)
    };
  }
}
