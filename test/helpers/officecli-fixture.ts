/**
 * OfficeCLI-backed fixtures for integration suites. The adapter is used
 * directly; residents are always closed so later rename/replace operations
 * are not blocked by file locks (§75 Windows notes).
 */

import { join } from "node:path";
import { OfficeCliAdapter } from "../../src/agent/officecli/officecli-adapter.js";

export interface OfficeCliFixture {
  adapter: OfficeCliAdapter;
  available: boolean;
  pptx: (dir: string, name?: string) => Promise<string>;
  docx: (dir: string, name?: string) => Promise<string>;
  xlsx: (dir: string, name?: string) => Promise<string>;
}

export async function createOfficeCliFixture(): Promise<OfficeCliFixture> {
  const adapter = new OfficeCliAdapter({ timeoutMs: 90_000 });
  const available = await adapter.version_().then(() => true).catch(() => false);

  async function create(dir: string, name: string, items: unknown[]): Promise<string> {
    const file = join(dir, name);
    await adapter.run(["create", file, "--json"]).catch(() => undefined);
    if (items.length > 0) {
      await adapter.runBatchStandalone(file, items);
    }
    // create/batch may leave a resident holding the file — always release.
    await adapter.close(file).catch(() => undefined);
    return file;
  }

  return {
    adapter,
    available,
    pptx: (dir, name = "fixture.pptx") =>
      create(dir, name, [
        { command: "add", parent: "/", type: "slide" },
        { command: "add", parent: "/slide[1]", type: "shape", props: { text: "Original Title", x: "1cm", y: "1cm" } },
        { command: "add", parent: "/slide[1]", type: "shape", props: { text: "Body content", x: "1cm", y: "4cm" } }
      ]),
    docx: (dir, name = "fixture.docx") =>
      create(dir, name, [
        { command: "add", parent: "/body", type: "paragraph", props: { text: "First paragraph" } },
        { command: "add", parent: "/body", type: "paragraph", props: { text: "Second paragraph" } }
      ]),
    xlsx: (dir, name = "fixture.xlsx") =>
      create(dir, name, [
        { command: "add", parent: "/", type: "sheet", props: { name: "Data" } },
        { command: "set", path: "/Data/A1", props: { value: "Header" } },
        { command: "set", path: "/Data/B1", props: { value: "42" } }
      ])
  };
}
