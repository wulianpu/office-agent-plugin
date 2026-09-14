// Temporary gate-verification fixture (issue #10): intentionally red.
import { describe, expect, it } from "vitest";
describe("gate check (issue #10)", () => {
  it("intentionally failing — merge MUST be blocked", () => {
    expect(1).toBe(2);
  });
});
