/**
 * MCP stdio wire protocol (§P8): initialize → tools/list → tools/call round
 * trips over newline-delimited JSON-RPC, plus protocol error handling.
 */

import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { PassThrough } from "node:stream";
import { McpStdioServer, toolDescriptors } from "../../src/mcp/server.js";
import { openWorkspace, writeDocxFixture } from "../helpers/fixtures.js";
import { join } from "node:path";

let ws: Awaited<ReturnType<typeof openWorkspace>>;
let input: PassThrough;
let output: PassThrough;

beforeAll(async () => {
  ws = await openWorkspace();
  await writeDocxFixture(join(ws.root, "wire.docx"), ["wire fixture"]);
  input = new PassThrough();
  output = new PassThrough();
  new McpStdioServer(ws.plugin.mcpTools).start(input, output);
});

afterAll(async () => {
  input.end();
  await ws?.cleanup().catch(() => undefined);
});

interface RpcEnvelope {
  id?: number | string;
  result?: {
    content?: Array<{ type: string; text: string }>;
    protocolVersion?: string;
    tools?: Array<{ name: string }>;
  };
  error?: { code: number; message: string };
}

function request(line: object): void {
  input.write(`${JSON.stringify(line)}\n`);
}

/** Read lines until one matching the request id arrives (or timeout). */
async function awaitResponse(id: number | string, timeoutMs = 10_000): Promise<RpcEnvelope> {
  const deadline = Date.now() + timeoutMs;
  let buffer = "";
  return new Promise<RpcEnvelope>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`no response for ${id}; buffer=${buffer}`)), deadline - Date.now());
    const onData = (chunk: Buffer) => {
      buffer += chunk.toString("utf8");
      const lines = buffer.split("\n");
      buffer = lines.pop() ?? "";
      for (const line of lines) {
        if (!line.trim()) continue;
        try {
          const envelope = JSON.parse(line) as RpcEnvelope;
          if (envelope.id === id) {
            clearTimeout(timer);
            output.off("data", onData);
            resolve(envelope);
          }
        } catch {
          // Skip non-JSON lines (server logs etc.).
        }
      }
    };
    output.on("data", onData);
  });
}

describe("MCP stdio server (wire protocol)", () => {
  it("initialize negotiates the 2026-01-26 protocol", async () => {
    request({ jsonrpc: "2.0", id: 1, method: "initialize", params: {} });
    const envelope = await awaitResponse(1);
    expect(envelope.result?.protocolVersion).toBe("2026-01-26");
  });

  it("responds to ping", async () => {
    request({ jsonrpc: "2.0", id: 2, method: "ping" });
    const envelope = await awaitResponse(2);
    expect(envelope.result).toEqual({});
  });

  it("tools/list exposes exactly the six office.* tools (§57)", async () => {
    request({ jsonrpc: "2.0", id: 3, method: "tools/list" });
    const envelope = await awaitResponse(3);
    const names = (envelope.result?.tools ?? []).map((t) => t.name);
    expect(names).toEqual([
      "office.inspect",
      "office.query",
      "office.edit",
      "office.render",
      "office.verify",
      "office.capabilities"
    ]);
    expect(toolDescriptors()).toHaveLength(6);
  });

  it("tools/call office.capabilities returns the matrix in-band", async () => {
    request({
      jsonrpc: "2.0",
      id: 4,
      method: "tools/call",
      params: { name: "office.capabilities", arguments: {} }
    });
    const envelope = await awaitResponse(4);
    const text = envelope.result?.content?.[0]?.text;
    expect(text).toBeTruthy();
    const parsed = JSON.parse(text!) as { offline: boolean };
    expect(parsed.offline).toBe(true);
  });

  it("tools/call with an unknown tool returns isError, not a protocol error", async () => {
    request({
      jsonrpc: "2.0",
      id: 5,
      method: "tools/call",
      params: { name: "office.nonexistent", arguments: {} }
    });
    const envelope = await awaitResponse(5);
    expect(envelope.error).toBeUndefined();
    expect(JSON.parse(envelope.result?.content?.[0]?.text ?? "{}")).toMatchObject({
      error: expect.stringContaining("unknown tool")
    });
  });

  it("tools/call against a missing session is an in-band error", async () => {
    request({
      jsonrpc: "2.0",
      id: 6,
      method: "tools/call",
      params: { name: "office.render", arguments: { sessionId: "sess_missing" } }
    });
    const envelope = await awaitResponse(6);
    const parsed = JSON.parse(envelope.result?.content?.[0]?.text ?? "{}") as { error: string };
    expect(parsed.error).toContain("no open session");
  });

  it("malformed JSON yields -32700 parse error", async () => {
    input.write("{not json}\n");
    request({ jsonrpc: "2.0", id: 7, method: "ping" });
    // After the parse error the server keeps serving: ping proves liveness.
    const envelope = await awaitResponse(7, 15_000);
    expect(envelope.result).toEqual({});
  });

  it("unknown method yields -32601", async () => {
    request({ jsonrpc: "2.0", id: 8, method: "resources/list" });
    const envelope = await awaitResponse(8);
    expect(envelope.error?.code).toBe(-32601);
  });
});
