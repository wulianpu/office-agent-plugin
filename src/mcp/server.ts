/**
 * MCP stdio server (§P8, §85): newline-delimited JSON-RPC 2.0. Implements the
 * 2026-01-26 protocol surface the plugin needs: initialize, tools/list,
 * tools/call, ping. The tool core is transport-agnostic; this is the wire.
 */

import { createInterface } from "node:readline";
import type { OfficeMcpTools } from "./tools.js";
import type { OfficeRuntimeService } from "../runtime/service/office-runtime-service.js";

const PROTOCOL_VERSION = "2026-01-26";

interface JsonRpcRequest {
  jsonrpc: "2.0";
  id?: number | string | null;
  method: string;
  params?: Record<string, unknown>;
}

interface JsonRpcNotification {
  jsonrpc: "2.0";
  method: string;
  params?: Record<string, unknown>;
}

export interface McpServerOptions {
  /** Host-side session registry: map documentId → sessionId before serving. */
  resolveSession?: (documentId: string) => string | undefined;
}

export function toolDescriptors() {
  return [
    {
      name: "office.inspect",
      description: "Inspect a document node (defaults to root) for the active session.",
      inputSchema: {
        type: "object",
        properties: {
          documentId: { type: "string" },
          sessionId: { type: "string" },
          path: { type: "string", description: "Logical document path, e.g. / or /slide[1]" }
        },
        required: ["sessionId"]
      }
    },
    {
      name: "office.query",
      description: "Query document elements with CSS-like selectors.",
      inputSchema: {
        type: "object",
        properties: {
          sessionId: { type: "string" },
          selector: { type: "string" }
        },
        required: ["sessionId", "selector"]
      }
    },
    {
      name: "office.edit",
      description: "Execute idempotent mutating batch items against the session candidate.",
      inputSchema: {
        type: "object",
        properties: {
          sessionId: { type: "string" },
          intent: { type: "string" },
          items: { type: "array" },
          scope: { type: "object" },
          idempotencyKey: { type: "string" }
        },
        required: ["sessionId", "intent", "items", "idempotencyKey"]
      }
    },
    {
      name: "office.render",
      description: "Render a bounded outline/text view of the document.",
      inputSchema: {
        type: "object",
        properties: { sessionId: { type: "string" }, mode: { type: "string", enum: ["text", "outline"] } },
        required: ["sessionId"]
      }
    },
    {
      name: "office.verify",
      description: "Fetch the hash-bound verification report for a candidate.",
      inputSchema: {
        type: "object",
        properties: { sessionId: { type: "string" }, candidateId: { type: "string" } },
        required: ["sessionId"]
      }
    },
    {
      name: "office.capabilities",
      description: "Capability matrix: per-format editor/agent/verification status.",
      inputSchema: { type: "object", properties: {} }
    }
  ];
}

export class McpStdioServer {
  constructor(
    private readonly tools: OfficeMcpTools,
    private readonly options: McpServerOptions = {}
  ) {}

  start(input: NodeJS.ReadableStream = process.stdin, output: NodeJS.WritableStream = process.stdout): void {
    const rl = createInterface({ input });
    rl.on("line", (line) => {
      const trimmed = line.trim();
      if (!trimmed) return;
      let message: JsonRpcRequest | JsonRpcNotification;
      try {
        message = JSON.parse(trimmed);
      } catch {
        this.respondError(null, -32700, "Parse error", output);
        return;
      }
      void this.handle(message, output);
    });
  }

  private async handle(
    message: JsonRpcRequest | JsonRpcNotification,
    output: NodeJS.WritableStream
  ): Promise<void> {
    const id: number | string | null | undefined =
      "id" in message ? (message.id as number | string | null | undefined) : undefined;
    const isRequest = id !== undefined && id !== null;
    switch (message.method) {
      case "initialize":
        if (isRequest) {
          this.respondOk(id, {
            protocolVersion: PROTOCOL_VERSION,
            capabilities: { tools: { listChanged: false } },
            serverInfo: { name: "office-plugin", version: "3.0.0" }
          }, output);
        }
        return;
      case "notifications/initialized":
      case "initialized":
        return;
      case "ping":
        if (isRequest) this.respondOk(id, {}, output);
        return;
      case "tools/list":
        if (isRequest) this.respondOk(id, { tools: toolDescriptors() }, output);
        return;
      case "tools/call": {
        if (!isRequest) return;
        const name = String((message.params as { name?: string })?.name ?? "");
        const args = ((message.params as { arguments?: Record<string, unknown> })?.arguments ?? {}) as never;
        try {
          const result = await this.callTool(name, args);
          this.respondOk(
            id,
            { content: [{ type: "text", text: JSON.stringify(result) }] },
            output
          );
        } catch (error) {
          // Tool errors are in-band results, not protocol errors.
          this.respondOk(
            id,
            {
              content: [{ type: "text", text: JSON.stringify({ error: String((error as Error)?.message ?? error) }) }],
              isError: true
            },
            output
          );
        }
        return;
      }
      default:
        if (isRequest) this.respondError(id, -32601, `Method not found: ${message.method}`, output);
    }
  }

  private callTool(name: string, args: never): Promise<unknown> | unknown {
    switch (name) {
      case "office.inspect":
        return this.tools.inspect(args);
      case "office.query":
        return this.tools.query(args);
      case "office.edit":
        return this.tools.edit(args);
      case "office.render":
        return this.tools.render(args);
      case "office.verify":
        return this.tools.verify(args);
      case "office.capabilities":
        return this.tools.capabilities();
      default:
        throw new Error(`unknown tool: ${name}`);
    }
  }

  private respondOk(id: number | string | null, result: unknown, output: NodeJS.WritableStream): void {
    output.write(`${JSON.stringify({ jsonrpc: "2.0", id, result })}\n`);
  }

  private respondError(
    id: number | string | null,
    code: number,
    message: string,
    output: NodeJS.WritableStream
  ): void {
    output.write(`${JSON.stringify({ jsonrpc: "2.0", id, error: { code, message } })}\n`);
  }
}
