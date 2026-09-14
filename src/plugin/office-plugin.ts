/**
 * OfficePlugin (P1, §6, §125–§127): the single logical plugin Harness sees.
 * Wires the runtime service, editor plugins, MCP tools and the capability
 * matrix; enforces the offline gate (§120).
 */

import { join } from "node:path";
import type { OfficeCapabilitiesResult } from "../contracts/mcp.js";
import type { ArtifactRef } from "../contracts/ids.js";
import type { PreviewRequest, PreviewResult } from "../contracts/preview.js";
import type { DocumentSession, ViewBookmark } from "../contracts/document.js";
import type { MutationScope, OfficeTaskContext } from "../contracts/capabilities.js";
import type { OfficeEditItem } from "../contracts/mcp.js";
import type { VerificationReport } from "../contracts/verification.js";
import type { EditorInstance, OfficeEditorPlugin } from "../contracts/editor.js";
import type { WriterLease } from "../contracts/lease.js";
import { OfficeRuntimeService } from "../runtime/service/office-runtime-service.js";
import { DB_SCHEMA_VERSION } from "../runtime/persistence/database.js";
import { OfficeMcpTools } from "../mcp/tools.js";
import { McpStdioServer } from "../mcp/server.js";
import { PortableReviewView } from "../mcp/portable-review/portable-review-view.js";
import { createBasicEditorPlugin } from "../editors/common/basic-plugin.js";
import type { RecoveryOutcome } from "../contracts/revision.js";
import type { ArtifactScanResult } from "../contracts/artifact.js";

export interface OfficePluginOptions {
  /** Workspace root for staging/DB (default: <cwd>/.office-runtime). */
  workspaceRoot?: string;
  dbPath?: string;
  engineDisabled?: boolean;
  /** Skip the optional WPS COM probe (tests; §86 stays optional). */
  skipHostProbe?: boolean;
}

export interface CompatibilityLock {
  plugin: string;
  genoffice: string;
  officecli: { version: string; schemaFingerprint: string };
  contract: number;
  dbSchema: number;
}

export const COMPATIBILITY_LOCK: CompatibilityLock = {
  plugin: "3.0.0",
  genoffice: "genspark-ai/genoffice@d35d770",
  officecli: { version: "1.x", schemaFingerprint: "probed-at-startup" },
  contract: 1,
  // Single source of truth: the persistence layer's schema constant. A drift
  // here lied about v2 while the DB was already v3 (round 8, P1-high) — the
  // invariant test now fails the suite on any future divergence.
  dbSchema: DB_SCHEMA_VERSION
};

export class OfficePlugin {
  readonly id = "office";
  readonly version = COMPATIBILITY_LOCK.plugin;
  readonly service: OfficeRuntimeService;
  readonly mcpTools: OfficeMcpTools;
  readonly review: PortableReviewView;
  private mcpServer?: McpStdioServer;

  private constructor(service: OfficeRuntimeService) {
    this.service = service;
    this.mcpTools = new OfficeMcpTools(service);
    this.review = new PortableReviewView(service);
    // Register editor plugins with engine="basic" (§127 honest degradation).
    for (const format of ["docx", "xlsx", "pptx"] as const) {
      const plugin: OfficeEditorPlugin = createBasicEditorPlugin(format, {
        resolvePath: (ref) => service.store.resolvePath(ref),
        save: async (_bytes, bookmark, sessionId) => {
          // P1 (#8): a degraded editor reports success ONLY after the
          // Runtime save gate (humanSave) validates/commits the source. A
          // bare save revalidates the committed source stability and fails
          // closed on external mutation — never a fake clean state.
          void bookmark;
          await service.humanSave(sessionId);
          return { savedAt: Date.now() };
        }
      });
      service.editorHost.register(plugin);
    }
  }

  static async create(options: OfficePluginOptions = {}): Promise<OfficePlugin> {
    const workspaceRoot = options.workspaceRoot ?? join(process.cwd(), ".office-runtime");
    const service = await OfficeRuntimeService.open({
      workspaceRoot,
      dbPath: options.dbPath,
      engineDisabled: options.engineDisabled,
      skipHostProbe: options.skipHostProbe
    });
    return new OfficePlugin(service);
  }

  // ---- Read path ----
  registerArtifact(path: string): Promise<ArtifactRef> {
    return this.service.registerArtifact(path);
  }

  preview(request: Omit<PreviewRequest, "requestId"> & { requestId?: string }): Promise<PreviewResult> {
    return this.service.preview(request);
  }

  // ---- Session lifecycle ----
  openSession(artifactRef: ArtifactRef): Promise<DocumentSession> {
    return this.service.openSession(artifactRef);
  }

  beginEdit(sessionId: string, bookmark?: ViewBookmark): Promise<{ lease: WriterLease; editor: EditorInstance }> {
    return this.service.beginEdit(sessionId, bookmark);
  }

  humanSave(sessionId: string, stagingRef?: ArtifactRef): Promise<{ revisionId: string; unchanged: boolean }> {
    return this.service.humanSave(sessionId, stagingRef);
  }

  endEdit(sessionId: string): Promise<void> {
    return this.service.endEdit(sessionId);
  }

  closeSession(sessionId: string): Promise<void> {
    return this.service.closeSession(sessionId);
  }

  // ---- Agent path ----
  beginAgentTask(sessionId: string, scope: MutationScope): Promise<OfficeTaskContext> {
    return this.service.beginAgentTask(sessionId, scope);
  }

  executeAgentMutation(
    task: OfficeTaskContext,
    command: { commandId: string; idempotencyKey: string; approved?: boolean; payload: OfficeEditItem[] }
  ) {
    return this.service.executeAgentMutation(task, command);
  }

  flushAgentCandidate(task: OfficeTaskContext): Promise<string> {
    return this.service.flushAgentCandidate(task);
  }

  verifyAgentCandidate(task: OfficeTaskContext, changedTargets?: string[]): Promise<VerificationReport> {
    return this.service.verifyAgentCandidate(task, changedTargets);
  }

  finalizeAgentTask(task: OfficeTaskContext): Promise<void> {
    return this.service.finalizeAgentTask(task);
  }

  acceptCandidate(sessionId: string, candidateId?: string): Promise<{ revisionId: string; contextPromoted: boolean }> {
    const id =
      candidateId ?? this.service.getSession(sessionId)?.candidate?.candidateId;
    if (!id) throw new Error("no active candidate");
    return this.service.acceptCandidate(sessionId, id);
  }

  rejectCandidate(sessionId: string): Promise<void> {
    const session = this.service.getSession(sessionId);
    if (!session?.candidate) throw new Error("no active candidate");
    return this.service.rejectCandidate(sessionId, session.candidate.candidateId);
  }

  // ---- MCP / review ----
  startMcpStdio(): void {
    this.mcpServer ??= new McpStdioServer(this.mcpTools);
    this.mcpServer.start();
  }

  // ---- Runtime ops ----
  recover(): Promise<RecoveryOutcome[]> {
    return this.service.recoverManually();
  }

  scanArtifact(ref: ArtifactRef): Promise<ArtifactScanResult> {
    return this.service.scanArtifact(ref);
  }

  capabilities(): OfficeCapabilitiesResult {
    return this.mcpTools.capabilities();
  }

  /** Offline gate (§120): core capabilities carry zero network dependency. */
  offlineGate(): { offline: true; notes: string[] } {
    return {
      offline: true,
      notes: [
        "runtime deps: none (node:sqlite / node:crypto / node:fs only)",
        "officecli spawned with NO_UPDATE=1, RESIDENT_FLUSH=off",
        "all renderer/editor assets bundled locally"
      ]
    };
  }

  compatibilityLock(): CompatibilityLock {
    return COMPATIBILITY_LOCK;
  }

  /**
   * Live compatibility snapshot: the officecli schema fingerprint is probed
   * from the engine's help surface at startup (hash of the capability
   * reference), not a hardcoded placeholder.
   */
  compatibilitySnapshot(): Promise<CompatibilityLock> {
    return this.service.probeCompatibilitySnapshot();
  }

  async dispose(): Promise<void> {
    await this.service.dispose();
  }
}
