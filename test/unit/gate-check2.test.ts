// Temporary gate-verification fixture (issue #10, round 2): intentionally red.
import { describe, expect, it } from "vitest";
describe("gate check 2 (issue #10)", () => {
  it("intentionally failing — merge MUST be blocked", () => {
    expect(2).toBe(3);
  });
});
