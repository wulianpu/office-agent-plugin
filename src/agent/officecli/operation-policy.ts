/**
 * OperationPolicyEngine (§124, §121): OfficeCLI *supporting* an operation ≠
 * the agent being allowed to run it. Risk ladder Low → Critical → Denied.
 */

import type { OfficeEditItem } from "../../contracts/mcp.js";
import type { MutationScope, OperationPolicy } from "../../contracts/capabilities.js";

/** Commands that never touch macros/OLE execution agents may issue. */
const DENIED_COMMANDS = new Set(["run", "exec", "macro", "activate", "watch", "unwatch"]);

/** Destructive verbs gated behind MutationScope.destructiveAllowed. */
const DESTRUCTIVE_COMMANDS = new Set(["remove", "raw-set"]);

/** Raw OOXML is Critical (§124) and additionally gated by allowedParts. */
const CRITICAL_COMMANDS = new Set(["raw-set", "raw", "add-part"]);

export const DEFAULT_RISK_THRESHOLD: OperationRiskThreshold = "high";

export type OperationRiskThreshold = "low" | "medium" | "high" | "critical";

const RISK_RANK: Record<OperationRiskThreshold, number> = {
  low: 0,
  medium: 1,
  high: 2,
  critical: 3
};

export interface PolicyDecision extends OperationPolicy {
  item: OfficeEditItem;
}

export class OperationPolicyEngine {
  constructor(
    /** Operations at or above this risk require explicit approval (default high). */
    readonly approvalThreshold: OperationRiskThreshold = DEFAULT_RISK_THRESHOLD
  ) {}

  classify(item: OfficeEditItem): OperationPolicy {
    const command = item.command?.toLowerCase() ?? "";
    if (DENIED_COMMANDS.has(command)) {
      return { risk: "denied", requiresApproval: true, reason: `command '${command}' is never agent-executable (§121)` };
    }
    if (CRITICAL_COMMANDS.has(command)) {
      return { risk: "critical", requiresApproval: true, reason: "raw OOXML mutation is Critical risk" };
    }
    if (command === "remove") {
      return { risk: "high", requiresApproval: true, reason: "destructive removal" };
    }
    if (command === "move" || command === "swap") {
      return { risk: "medium", requiresApproval: false };
    }
    if (command === "add") {
      // Adding slides/sheets reshapes structure; adding text into existing
      // containers is lower risk but still structural → medium.
      return { risk: "medium", requiresApproval: false };
    }
    if (command === "set") {
      const props = item.props ?? {};
      const touchesStructure = "style" in props || "theme" in props || "master" in props;
      return {
        risk: touchesStructure ? "high" : "low",
        requiresApproval: false
      };
    }
    if (command === "import") {
      return { risk: "high", requiresApproval: true, reason: "bulk data import" };
    }
    return { risk: "medium", requiresApproval: false, reason: `unclassified command '${command}' defaults to medium` };
  }

  /** Combine the risk ladder with MutationScope fences (§60). */
  evaluate(item: OfficeEditItem, scope: MutationScope, approved: boolean): PolicyDecision {
    const policy = this.classify(item);
    if (policy.risk === "denied") return { ...policy, item };

    const reasons: string[] = [];
    if (policy.reason) reasons.push(policy.reason);

    // Scope: destructive commands need destructiveAllowed.
    if (DESTRUCTIVE_COMMANDS.has(item.command?.toLowerCase() ?? "") && !scope.destructiveAllowed) {
      return {
        item,
        risk: "denied",
        requiresApproval: true,
        reason: `command '${item.command}' is destructive but MutationScope.destructiveAllowed=false`
      };
    }

    // Scope: target fencing. When allowedTargets is set, every path/parent/
    // selector must fall under one of the allowed prefixes.
    if (scope.allowedTargets && scope.allowedTargets.length > 0) {
      const targets = [item.path, item.parent, item.selector, item.to, item.path2].filter(
        (t): t is string => typeof t === "string" && t.length > 0
      );
      for (const target of targets) {
        if (!scope.allowedTargets.some((allowed) => target === allowed || target.startsWith(`${allowed}/`) || target.startsWith(`${allowed}[`))) {
          return {
            item,
            risk: "denied",
            requiresApproval: true,
            reason: `target '${target}' outside MutationScope.allowedTargets`
          };
        }
      }
    }

    // Scope: raw part fencing.
    if (scope.allowedParts && scope.allowedParts.length > 0 && item.part) {
      if (!scope.allowedParts.includes(item.part)) {
        return {
          item,
          risk: "denied",
          requiresApproval: true,
          reason: `part '${item.part}' outside MutationScope.allowedParts`
        };
      }
    }

    const requiresApproval =
      policy.risk !== "low" &&
      policy.risk !== "medium" &&
      RANK(policy.risk) >= RANK(this.approvalThreshold) &&
      !approved;

    return { ...policy, requiresApproval, reason: reasons.join("; ") || undefined, item };
  }
}

function RANK(risk: "low" | "medium" | "high" | "critical"): number {
  return RISK_RANK[risk];
}
