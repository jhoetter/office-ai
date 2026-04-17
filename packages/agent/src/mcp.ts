/**
 * OfficeAI Model Context Protocol server.
 *
 * Exposes the headless `DocxAgent` as a set of MCP tools. Sessions live
 * in-process; each successful `docx_load` mints an opaque `handle` that
 * other tools accept. Handles are intentionally short-lived (process
 * lifetime) — they are NOT persistent identifiers.
 *
 * Transport-agnostic: `runMcpStdioServer()` wires up `StdioServerTransport`
 * for the published binary, but tests use `InMemoryTransport` directly via
 * the exported `createMcpServer()` factory.
 */

import { readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { randomUUID } from "node:crypto";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { DocxAgent } from "@officeai/docx";
import { diffSnapshots, inspectSnapshot, snapshotToJsonProjection } from "./cli.js";

const sessions = new Map<string, DocxAgent>();
const sessionPaths = new Map<string, string>();

/** Test hook: drop in-memory state between test cases. */
export function __resetMcpSessionsForTests(): void {
  sessions.clear();
  sessionPaths.clear();
}

function lookupAgent(handle: string): DocxAgent {
  const agent = sessions.get(handle);
  if (!agent) {
    throw new Error(`Unknown DOCX handle: "${handle}". Call docx_load first.`);
  }
  return agent;
}

function ok(payload: unknown): {
  content: Array<{ type: "text"; text: string }>;
  structuredContent: Record<string, unknown>;
} {
  const text = JSON.stringify(payload, null, 2);
  return {
    content: [{ type: "text", text }],
    structuredContent: payload as Record<string, unknown>,
  };
}

function fail(message: string): {
  isError: true;
  content: Array<{ type: "text"; text: string }>;
} {
  return {
    isError: true,
    content: [{ type: "text", text: message }],
  };
}

/**
 * Build a fresh MCP server with all OfficeAI tools registered. Exposed for
 * tests so they can wire it to an in-memory transport pair without touching
 * stdio.
 */
export function createMcpServer(): McpServer {
  const server = new McpServer(
    { name: "officeai", version: "0.1.0" },
    {
      capabilities: { tools: {} },
      instructions:
        "Tools for parsing, inspecting, and editing OOXML DOCX files. Always start with `docx_load` to obtain a handle, then pass that handle to other tools. Call `docx_save` (or pass `out_path`) to persist edits.",
    }
  );

  // ── docx_load ─────────────────────────────────────────────────────────
  server.registerTool(
    "docx_load",
    {
      description: "Load a .docx file from disk. Returns an opaque `handle` to use with subsequent tools.",
      inputSchema: {
        path: z.string().describe("Absolute or workspace-relative path to a .docx file."),
      },
    },
    async ({ path }) => {
      try {
        const buf = await readFile(resolve(path));
        const agent = await DocxAgent.fromBuffer(buf);
        const handle = randomUUID();
        sessions.set(handle, agent);
        sessionPaths.set(handle, resolve(path));
        return ok({ handle, path: resolve(path), summary: inspectSnapshot(agent.getSnapshot()) });
      } catch (err) {
        return fail(`docx_load failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_save ─────────────────────────────────────────────────────────
  server.registerTool(
    "docx_save",
    {
      description:
        "Serialize the current snapshot back to disk. Defaults to the path passed to docx_load; pass `out_path` to write elsewhere.",
      inputSchema: {
        handle: z.string().describe("Handle returned by docx_load."),
        out_path: z
          .string()
          .optional()
          .describe("Optional output path. Defaults to the original path passed to docx_load."),
      },
    },
    async ({ handle, out_path }) => {
      try {
        const agent = lookupAgent(handle);
        agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
        const target = out_path ? resolve(out_path) : sessionPaths.get(handle);
        if (!target) {
          return fail(`docx_save: no path known for handle "${handle}". Pass out_path explicitly.`);
        }
        const buf = Buffer.from(await agent.exportFile());
        await writeFile(target, buf);
        return ok({ wrote: target, bytes: buf.byteLength, revision: agent.getSnapshot().revision });
      } catch (err) {
        return fail(`docx_save failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_inspect ──────────────────────────────────────────────────────
  server.registerTool(
    "docx_inspect",
    {
      description: "Return a structural summary (paragraphs, tables, comments, parts).",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupAgent(handle);
        return ok(inspectSnapshot(agent.getSnapshot()));
      } catch (err) {
        return fail(`docx_inspect failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_get_text ─────────────────────────────────────────────────────
  server.registerTool(
    "docx_get_text",
    {
      description: "Return the document content as Markdown (default), structured JSON, or plain text.",
      inputSchema: {
        handle: z.string(),
        format: z.enum(["markdown", "json", "text"]).optional().default("markdown"),
      },
    },
    async ({ handle, format }) => {
      try {
        const agent = lookupAgent(handle);
        const fmt = format ?? "markdown";
        switch (fmt) {
          case "markdown":
            return ok({ format: fmt, content: agent.toMarkdown() });
          case "json":
            return ok(snapshotToJsonProjection(agent.getSnapshot()));
          case "text": {
            const lines: string[] = [];
            for (const b of agent.getSnapshot().root.body) {
              if (b.kind === "paragraph") {
                lines.push(
                  b.children
                    .map((c) =>
                      c.kind === "run"
                        ? c.children.map((g) => (g.kind === "text" ? g.text : "")).join("")
                        : ""
                    )
                    .join("")
                );
              }
            }
            return ok({ format: fmt, content: lines.join("\n") });
          }
          default: {
            const _exhaustive: never = fmt;
            void _exhaustive;
            return fail(`unknown format: ${String(format)}`);
          }
        }
      } catch (err) {
        return fail(`docx_get_text failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_search ───────────────────────────────────────────────────────
  server.registerTool(
    "docx_search",
    {
      description: "Search the document body for text. Optional case-sensitive and regex flags.",
      inputSchema: {
        handle: z.string(),
        query: z.string().min(1),
        case_sensitive: z.boolean().optional().default(false),
        regex: z.boolean().optional().default(false),
      },
    },
    async ({ handle, query, case_sensitive, regex }) => {
      try {
        const agent = lookupAgent(handle);
        const results = agent.search({
          query,
          caseSensitive: case_sensitive ?? false,
          regex: regex ?? false,
        });
        return ok({ matches: results });
      } catch (err) {
        return fail(`docx_search failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_apply_command ────────────────────────────────────────────────
  server.registerTool(
    "docx_apply_command",
    {
      description:
        "Apply a docx command (e.g. `docx:insert-text`, `docx:resolve-comment`). Pass an arbitrary payload object — schemas live in `@officeai/docx/commands/payloads`.",
      inputSchema: {
        handle: z.string(),
        type: z.string().describe('Command type, e.g. "docx:insert-text"'),
        payload: z.record(z.string(), z.unknown()).describe("Command payload."),
        source: z.enum(["agent", "human", "system"]).optional().default("agent"),
        agent_id: z.string().optional().default("officeai-mcp"),
        auto_approve: z
          .boolean()
          .optional()
          .default(true)
          .describe(
            "When true (default) any pending mutation produced by an `agent` source is immediately approved. Set false to leave it pending in the bus for downstream review."
          ),
      },
    },
    async ({ handle, type, payload, source, agent_id, auto_approve }) => {
      try {
        const agent = lookupAgent(handle);
        const mutation = await agent.applyCommand({
          type,
          payload,
          source: source ?? "agent",
          ...((source ?? "agent") === "agent" ? { agentId: agent_id ?? "officeai-mcp" } : {}),
        });
        if ((auto_approve ?? true) && mutation.status === "pending") {
          agent.approveMutation(mutation.id);
        }
        return ok({
          mutation: {
            id: mutation.id,
            status: agent.getPendingMutations().some((m) => m.id === mutation.id)
              ? "pending"
              : mutation.status,
            ...(mutation.rejection ? { rejection: mutation.rejection } : {}),
          },
          revision: agent.getSnapshot().revision,
        });
      } catch (err) {
        return fail(`docx_apply_command failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_diff ─────────────────────────────────────────────────────────
  server.registerTool(
    "docx_diff",
    {
      description:
        "Diff two loaded handles, OR diff a handle against the file on disk it was loaded from. Pass either {before, after} (two handles) or {handle, against: 'disk'}.",
      inputSchema: {
        before: z.string().optional(),
        after: z.string().optional(),
        handle: z.string().optional(),
        against: z.enum(["disk"]).optional(),
      },
    },
    async ({ before, after, handle, against }) => {
      try {
        if (before && after) {
          const a = lookupAgent(before);
          const b = lookupAgent(after);
          return ok(diffSnapshots(a.getSnapshot(), b.getSnapshot()));
        }
        if (handle && against === "disk") {
          const agent = lookupAgent(handle);
          const path = sessionPaths.get(handle);
          if (!path) return fail(`docx_diff: no on-disk path for handle "${handle}".`);
          const buf = await readFile(path);
          const baseline = await DocxAgent.fromBuffer(buf);
          return ok(diffSnapshots(baseline.getSnapshot(), agent.getSnapshot()));
        }
        return fail("docx_diff: pass either {before, after} (two handles) or {handle, against: 'disk'}.");
      } catch (err) {
        return fail(`docx_diff failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  return server;
}

/**
 * Run the MCP server over stdio. This is the entry point used by
 * `office-agent mcp` — it never returns under normal operation; the SDK's
 * stdio transport keeps the process alive until the parent closes stdin.
 */
export async function runMcpStdioServer(): Promise<void> {
  const server = createMcpServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  // Keep the event loop alive until stdin closes.
  await new Promise<void>((resolveP) => {
    process.stdin.on("close", () => resolveP());
  });
}
