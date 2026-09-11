/**
 * Format editor plugins (§44): one per format, engine="basic" until GenOffice
 * is vendored (§146). The capability matrix surfaces the engine + degradation
 * honestly (§127).
 */

import type {
  EditorBootstrapContext,
  EditorInstance,
  OfficeEditorPlugin
} from "../../contracts/editor.js";
import type { OfficeFormat } from "../../contracts/ids.js";
import { BasicEditorInstance, outlineLoaderFor, type BasicEditorSave } from "../common/basic-editor.js";

export interface BasicPluginDeps {
  resolvePath: (ref: string) => string;
  save: BasicEditorSave;
}

export function createBasicEditorPlugin(
  format: OfficeFormat,
  deps: BasicPluginDeps
): OfficeEditorPlugin {
  return {
    format,
    engine: "basic",
    async create(context: EditorBootstrapContext): Promise<EditorInstance> {
      const editor = new BasicEditorInstance(
        format,
        context,
        outlineLoaderFor(format),
        deps.resolvePath,
        deps.save
      );
      return editor;
    }
  };
}
