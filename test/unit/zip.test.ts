import { describe, expect, it } from "vitest";
import { writeFile, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { readZipEntry, readZipIndex, hasContentTypes } from "../../src/artifact/scanner/zip.js";
import { buildZip } from "../helpers/zip-builder.js";

describe("zip reader", () => {
  it("parses central directory of a stored archive", async () => {
    const dir = await mkdtemp(join(tmpdir(), "zip-test-"));
    const path = join(dir, "test.docx");
    await writeFile(
      path,
      buildZip([
        { name: "[Content_Types].xml", data: "<Types/>" },
        { name: "word/document.xml", data: "<doc>hello</doc>" }
      ])
    );
    try {
      const index = await readZipIndex(path);
      expect(index.entries).toHaveLength(2);
      expect(hasContentTypes(index)).toBe(true);
      const doc = await readZipEntry(path, index, "word/document.xml");
      expect(doc.toString("utf8")).toBe("<doc>hello</doc>");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("rejects non-zip files", async () => {
    const dir = await mkdtemp(join(tmpdir(), "zip-test-"));
    const path = join(dir, "plain.txt");
    await writeFile(path, "not a zip");
    try {
      await expect(readZipIndex(path)).rejects.toThrow();
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
