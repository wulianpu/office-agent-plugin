/**
 * Editor host + registry (§44–§47): Harness Core depends on the plugin surface
 * only. InProcessEditorHost is the first-version host (§46); an isolated host
 * can be added later without runtime changes (§47).
 */

import type {
  EditorBootstrapContext,
  EditorContainer,
  EditorHost,
  EditorInstance,
  OfficeEditorPlugin
} from "../../contracts/editor.js";
import type { OfficeFormat } from "../../contracts/ids.js";

export class InProcessEditorHost implements EditorHost {
  private readonly registry = new Map<OfficeFormat, OfficeEditorPlugin>();

  register(plugin: OfficeEditorPlugin): void {
    this.registry.set(plugin.format, plugin);
  }

  pluginFor(format: OfficeFormat): OfficeEditorPlugin | undefined {
    return this.registry.get(format);
  }

  formatsAvailable(): OfficeFormat[] {
    return [...this.registry.keys()];
  }

  async mount(
    plugin: OfficeEditorPlugin,
    context: EditorBootstrapContext
  ): Promise<EditorInstance> {
    const instance = await plugin.create(context);
    // §46: in-process host mounts into the provided container; headless
    // containers are accepted so the same contract works in tests/utility host.
    const container: EditorContainer = { kind: "headless" };
    await instance.mount(container);
    return instance;
  }
}

/** Headless container factory for non-DOM hosts (CLI, tests, utility process). */
export function headlessContainer(): EditorContainer {
  return { kind: "headless" };
}
