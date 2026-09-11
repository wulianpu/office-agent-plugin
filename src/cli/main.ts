/**
 * Demo CLI: walks the full Preview → Open → Agent Candidate → Verify →
 * Review → Accept lifecycle (§157–§158) against a real file.
 */

import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { OfficePlugin } from "../plugin/office-plugin.js";
import { newCommandId } from "../support/ids.js";

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const file = args[0];
  if (!file) {
    console.error("usage: office-runtime-demo <file.(docx|xlsx|pptx)>");
    process.exit(2);
  }

  const workspace = await mkdtemp(join(tmpdir(), "office-runtime-demo-"));
  const plugin = await OfficePlugin.create({ workspaceRoot: workspace });
  try {
    console.log("[capabilities]", JSON.stringify(plugin.capabilities(), null, 2));

    const ref = await plugin.registerArtifact(file);
    const preview = await plugin.preview({ artifactRef: ref, priority: "visible" });
    console.log(
      `[preview] format=%s elapsedMs=%d retried=%s outlineSummary=%s`,
      preview.model.format,
      preview.elapsedMs,
      preview.retried,
      preview.model.outline.kind
    );

    const session = await plugin.openSession(ref);
    console.log(
      "[open] sessionId=%s revision=%s hash=%s",
      session.sessionId,
      session.committedRevision.revisionId,
      session.committedRevision.contentHash.slice(0, 12)
    );

    if (plugin.service.isEngineAvailable()) {
      const task = await plugin.beginAgentTask(session.sessionId, {
        intent: "demo: set title emphasis",
        destructiveAllowed: false
      });
      await plugin.executeAgentMutation(task, {
        commandId: newCommandId(),
        idempotencyKey: "demo-1",
        payload:
          session.format === "pptx"
            ? [{ command: "set", path: "/slide[1]/shape[1]", props: { text: "Demo Title (agent)" } }]
            : [{ command: "set", path: "/body/paragraph[1]", props: { text: "Demo Title (agent)" } }]
      });
      const hash = await plugin.flushAgentCandidate(task);
      console.log("[agent] candidate=%s flushed=%s", task.candidateId, hash.slice(0, 12));
      const report = await plugin.verifyAgentCandidate(task);
      console.log("[verify] confidence=%s", report.confidence);
      await plugin.finalizeAgentTask(task);
      const review = await plugin.review.build(session.sessionId);
      console.log(
        "[review] candidateState=%s verification=%s",
        review.candidate.state,
        review.verification?.confidence ?? "none"
      );
      const accepted = await plugin.acceptCandidate(session.sessionId);
      console.log("[accept] revision=%s promoted=%s", accepted.revisionId, accepted.contextPromoted);
    } else {
      console.log("[agent] officecli unavailable — demonstrating read path only");
    }

    await plugin.closeSession(session.sessionId);
    console.log("[done]");
  } finally {
    await plugin.dispose();
    await rm(workspace, { recursive: true, force: true });
  }
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
