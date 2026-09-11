/**
 * BasicEditorInstance — contract-complete headless editors (§45). The visual
 * GenOffice runtime is vendored later (§146); until then these editors cover
 * the runtime-facing lifecycle (mount/activateEdit/save/reload/suspend/resume/
 * dispose) against an outline data model, and the capability matrix reports
 * them as degraded. Save delegates to the runtime's human-save callback so
 * revision/commit flows are fully exercised.
 */

import type {
  EditorBootstrapContext,
  EditorInstance,
  EditorSaveResult,
  EditorState,
  ReloadRequest,
  SelectionAnchor
} from "../../contracts/editor.js";
import type { OfficeFormat, SessionId } from "../../contracts/ids.js";
import type { ViewBookmark } from "../../contracts/document.js";
import { renderDocxOutline, renderPptxOutline, renderXlsxOutline } from "../../preview/outline-renderers.js";
import { newId } from "../../support/ids.js";

export interface BasicEditorSave {
  (bytes: undefined, bookmark: ViewBookmark | undefined): Promise<EditorSaveResult>;
}

export class BasicEditorInstance implements EditorInstance {
  readonly instanceId = newId("edt");
  state: EditorState = "detached";
  private model?: unknown;
  private bookmark?: ViewBookmark;
  private selection?: SelectionAnchor;
  private suspended = false;

  constructor(
    readonly format: OfficeFormat,
    private readonly bootstrap: EditorBootstrapContext,
    private readonly loadModel: (path: string) => Promise<unknown>,
    private readonly resolvePath: (ref: string) => string,
    private readonly saveImpl: BasicEditorSave
  ) {}

  async mount(): Promise<void> {
    this.state = "loading";
    this.model = await this.loadModel(this.resolvePath(this.bootstrap.artifactContext.artifactRef));
    this.bookmark = this.bootstrap.bookmark;
    this.state = this.bootstrap.readOnly ? "clean" : "clean";
  }

  /**
   * Edit activation (§13): the runtime has already performed promotion checks
   * and acquired the human lease by the time this is invoked.
   */
  async activateEdit(): Promise<void> {
    if (this.suspended) throw new Error("editor suspended");
    this.state = "clean"; // editable runtime active; dirty on first mutation
  }

  /** Simulated human mutation (used by tests/demo to drive the save path). */
  markDirty(): void {
    if (this.state === "clean") this.state = "dirty";
  }

  async save(): Promise<EditorSaveResult> {
    if (this.state === "detached" || this.state === "suspended") {
      throw new Error(`cannot save from state ${this.state}`);
    }
    this.state = "saving";
    try {
      const result = await this.saveImpl(undefined, this.bookmark);
      this.state = "clean";
      return result;
    } catch (error) {
      this.state = "error";
      throw error;
    }
  }

  async reload(input: ReloadRequest): Promise<void> {
    this.state = "loading";
    this.model = await this.loadModel(this.resolvePath(input.artifactContext.artifactRef));
    this.bookmark = input.bookmark;
    this.state = "clean";
  }

  async getSelection(): Promise<SelectionAnchor | null> {
    return this.selection ?? null;
  }

  setSelection(anchor: SelectionAnchor): void {
    this.selection = anchor;
  }

  async suspend(): Promise<void> {
    this.suspended = true;
    this.state = this.state === "dirty" ? "dirty" : "suspended";
  }

  async resume(): Promise<void> {
    this.suspended = false;
    this.state = "clean";
  }

  async dispose(): Promise<void> {
    this.model = undefined;
    this.state = "detached";
  }

  currentModel(): unknown {
    return this.model;
  }

  sessionId(): SessionId {
    return this.bootstrap.sessionId;
  }
}

/** Loader per format for the outline models. */
export function outlineLoaderFor(format: OfficeFormat): (path: string) => Promise<unknown> {
  switch (format) {
    case "docx":
      return renderDocxOutline;
    case "xlsx":
      return renderXlsxOutline;
    case "pptx":
      return renderPptxOutline;
  }
}
