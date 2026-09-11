/**
 * Phase 0 — Contracts: editor plane (Design Doc §44–§55).
 */

import type { OfficeFormat, RevisionId, SessionId } from "./ids.js";
import type { ArtifactContext } from "./artifact.js";
import type { ViewBookmark } from "./document.js";

/** §34: editor state machine, independent of session/candidate states. */
export type EditorState =
  | "detached"
  | "loading"
  | "clean"
  | "dirty"
  | "saving"
  | "suspended"
  | "error";

export interface EditorBootstrapContext {
  sessionId: SessionId;
  artifactContext: ArtifactContext;
  /** Bookmark captured before promotion so edit never jumps back to the start (§54). */
  bookmark?: ViewBookmark;
  readOnly: boolean;
}

export interface EditorSaveResult {
  /** Artifact content hash after save, when the save flushed bytes. */
  contentHash?: string;
  savedAt: number;
}

export interface ReloadRequest {
  artifactContext: ArtifactContext;
  bookmark?: ViewBookmark;
}

/**
 * §55: revision-aware selection anchors. Agents must resolve → probe → confirm
 * before mutating (Read → Probe → Mutate).
 */
export interface SelectionAnchor {
  revisionId: RevisionId;
  format: OfficeFormat;
  logicalPath?: string;
  stableId?: string;
  containerId?: string;
  fingerprint?: string;
}

/** §45: the runtime-facing editor instance surface. */
export interface EditorInstance {
  readonly instanceId: string;
  readonly format: OfficeFormat;
  state: EditorState;

  mount(container: EditorContainer): Promise<void>;
  activateEdit(): Promise<void>;
  save(): Promise<EditorSaveResult>;
  reload(input: ReloadRequest): Promise<void>;
  getSelection(): Promise<SelectionAnchor | null>;
  suspend(): Promise<void>;
  resume(): Promise<void>;
  dispose(): Promise<void>;
  /** Optional: hosts/tests use it to signal a first human mutation. */
  markDirty?(): void;
}

/**
 * Headless stand-in for a DOM container. The first version runs editors
 * in-process (§46 InProcessEditorHost); a DOM container adapter can be
 * layered later without touching the contract.
 */
export interface EditorContainer {
  readonly kind: "dom" | "headless";
  readonly handle?: unknown;
}

/** §44: Harness Core depends on this plugin surface, never on GenOffice APIs. */
export interface OfficeEditorPlugin {
  format: OfficeFormat;
  /** Engine identifier reported through the capability matrix (e.g. "genoffice", "basic"). */
  engine: string;
  create(context: EditorBootstrapContext): Promise<EditorInstance>;
}

/** §47: editor host abstraction — InProcess now, Isolated later without runtime changes. */
export interface EditorHost {
  mount(plugin: OfficeEditorPlugin, context: EditorBootstrapContext): Promise<EditorInstance>;
}

export interface EditorBinding {
  instanceId: string;
  plugin: OfficeEditorPlugin;
  host: EditorHost;
  boundAt: number;
}
