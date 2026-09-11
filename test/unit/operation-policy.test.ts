import { describe, expect, it } from "vitest";
import { OperationPolicyEngine } from "../../src/agent/officecli/operation-policy.js";

const scope = (over: Partial<{ allowedTargets: string[]; destructiveAllowed: boolean; allowedParts: string[] }> = {}) => ({
  intent: "test",
  destructiveAllowed: false,
  ...over
});

describe("OperationPolicyEngine (§124, §60)", () => {
  const engine = new OperationPolicyEngine();

  it("classifies text edits low and structural edits higher", () => {
    expect(engine.classify({ command: "set", path: "/slide[1]/shape[1]", props: { bold: "true" } }).risk).toBe("low");
    expect(engine.classify({ command: "add", parent: "/slide[1]", type: "shape" }).risk).toBe("medium");
    expect(engine.classify({ command: "remove", path: "/slide[2]" }).risk).toBe("high");
    expect(engine.classify({ command: "raw-set", part: "/ppt/slides/slide1.xml", content: "<xml/>" }).risk).toBe("critical");
  });

  it("denies macro/execution commands outright (§121)", () => {
    const decision = engine.evaluate({ command: "macro", path: "/" }, scope(), false);
    expect(decision.risk).toBe("denied");
  });

  it("gates destructive commands behind MutationScope.destructiveAllowed", () => {
    const item = { command: "remove", path: "/slide[3]/shape[2]" };
    expect(engine.evaluate(item, scope(), false).risk).toBe("denied");
    const allowed = engine.evaluate(item, scope({ destructiveAllowed: true }), false);
    expect(allowed.risk).toBe("high");
    expect(allowed.requiresApproval).toBe(true); // high ≥ threshold
  });

  it("fences targets to allowedTargets (§60)", () => {
    const item = { command: "set", path: "/slide[3]/shape[1]", props: { color: "#ff0000" } };
    expect(
      engine.evaluate(item, scope({ allowedTargets: ["/slide[3]"] }), false).risk
    ).toBe("low");
    const denied = engine.evaluate(item, scope({ allowedTargets: ["/slide[4]"] }), false);
    expect(denied.risk).toBe("denied");
    expect(denied.reason).toContain("allowedTargets");
  });

  it("fences raw parts to allowedParts", () => {
    const item = { command: "raw-set", part: "ppt/slides/slide3.xml", content: "<xml/>" };
    const rawScope = scope({ allowedParts: ["ppt/slides/slide3.xml"], destructiveAllowed: true });
    expect(engine.evaluate(item, rawScope, false).risk).toBe("critical");
    const deniedScope = scope({ allowedParts: ["ppt/theme/theme1.xml"], destructiveAllowed: true });
    const denied = engine.evaluate(item, deniedScope, false);
    expect(denied.risk).toBe("denied");
  });
});
