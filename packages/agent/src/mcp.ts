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
import { XlsxAgent, diffXlsxSnapshots } from "@officeai/xlsx";
import { PptxAgent } from "@officeai/pptx";
import { diffSnapshots, inspectSnapshot, snapshotToJsonProjection } from "./cli.js";
import { inspectXlsxSnapshot, xlsxRangeToJson } from "./cli-xlsx.js";
import {
  diffSnapshots as pptxDiffSnapshots,
  inspectSnapshot as pptxInspectSnapshot,
  snapshotToJsonProjection as pptxSnapshotToJsonProjection,
} from "./pptx-cli.js";

const sessions = new Map<string, DocxAgent>();
const sessionPaths = new Map<string, string>();
const xlsxSessions = new Map<string, XlsxAgent>();
const xlsxSessionPaths = new Map<string, string>();
const pptxSessions = new Map<string, PptxAgent>();
const pptxSessionPaths = new Map<string, string>();

/** Test hook: drop in-memory state between test cases. */
export function __resetMcpSessionsForTests(): void {
  sessions.clear();
  sessionPaths.clear();
  xlsxSessions.clear();
  xlsxSessionPaths.clear();
  pptxSessions.clear();
  pptxSessionPaths.clear();
}

function lookupAgent(handle: string): DocxAgent {
  const agent = sessions.get(handle);
  if (!agent) {
    throw new Error(`Unknown DOCX handle: "${handle}". Call docx_load first.`);
  }
  return agent;
}

function lookupXlsxAgent(handle: string): XlsxAgent {
  const agent = xlsxSessions.get(handle);
  if (!agent) {
    throw new Error(`Unknown XLSX handle: "${handle}". Call xlsx_load first.`);
  }
  return agent;
}

function lookupPptxAgent(handle: string): PptxAgent {
  const agent = pptxSessions.get(handle);
  if (!agent) {
    throw new Error(`Unknown PPTX handle: "${handle}". Call pptx_load first.`);
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
        "Tools for parsing, inspecting, and editing OOXML DOCX, XLSX and PPTX files. Always start with `docx_load` (DOCX), `xlsx_load` (XLSX), or `pptx_load` (PPTX) to obtain a handle, then pass that handle to other tools. Call `docx_save` / `xlsx_save` / `pptx_save` (or pass `out_path`) to persist edits.",
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
      description:
        "Return the document content as Markdown (default), structured JSON, or plain text. Pass `with_page_sections: true` to interleave `<!-- page N -->` anchors so the LLM can cite pages.",
      inputSchema: {
        handle: z.string(),
        format: z.enum(["markdown", "json", "text"]).optional().default("markdown"),
        with_page_sections: z
          .boolean()
          .optional()
          .default(false)
          .describe(
            "When true and format is markdown, prepend each page with a <!-- page N --> anchor + ## Page N heading."
          ),
      },
    },
    async ({ handle, format, with_page_sections }) => {
      try {
        const agent = lookupAgent(handle);
        const fmt = format ?? "markdown";
        switch (fmt) {
          case "markdown":
            return ok({
              format: fmt,
              content: agent.toMarkdown(with_page_sections ? { withPageSections: true } : undefined),
            });
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

  // ── docx_get_pages ────────────────────────────────────────────────────
  server.registerTool(
    "docx_get_pages",
    {
      description:
        "List the document's logical pages with their body-block range, the trigger that started each page (doc-start, page-break, last-rendered, section-break), and a short text preview. Page numbers are 1-based and global across the document.",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupAgent(handle);
        const pages = agent.getPages();
        return ok({ pages, total: pages.length });
      } catch (err) {
        return fail(`docx_get_pages failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_get_page_text ────────────────────────────────────────────────
  server.registerTool(
    "docx_get_page_text",
    {
      description:
        "Return the markdown (default) or plain-text projection of a single page. Pass `page` (1-based) and optionally `format`. Use docx_get_pages first to discover the available page numbers.",
      inputSchema: {
        handle: z.string(),
        page: z.number().int().positive(),
        format: z.enum(["markdown", "text"]).optional().default("markdown"),
      },
    },
    async ({ handle, page, format }) => {
      try {
        const agent = lookupAgent(handle);
        const fmt = format ?? "markdown";
        const pages = agent.getPages();
        const info = pages.find((p) => p.pageNumber === page);
        if (!info) {
          return fail(
            `docx_get_page_text: page ${page} out-of-range (document has ${pages.length} page${pages.length === 1 ? "" : "s"})`
          );
        }
        const content = fmt === "markdown" ? agent.getPageMarkdown(page) : agent.getPageText(page);
        if (content === null) {
          return fail(`docx_get_page_text: page ${page} out-of-range`);
        }
        return ok({
          pageNumber: page,
          startBlockIndex: info.startBlockIndex,
          endBlockIndex: info.endBlockIndex,
          format: fmt,
          content,
        });
      } catch (err) {
        return fail(`docx_get_page_text failed: ${err instanceof Error ? err.message : String(err)}`);
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

  // ── docx_list_pending ─────────────────────────────────────────────────
  server.registerTool(
    "docx_list_pending",
    {
      description:
        "List mutations that are still in `pending` state on the bus (typically agent-authored writes invoked with auto_approve=false). Each entry includes the mutation id, command type, source, and the agent id that submitted it.",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupAgent(handle);
        const pending = agent.getPendingMutations().map((m) => ({
          id: m.id,
          command: { type: m.command.type, source: m.command.source },
          ...(m.command.source === "agent" && "agentId" in m.command ? { agentId: m.command.agentId } : {}),
          revision: m.after.revision,
          status: m.status,
        }));
        return ok({ pending });
      } catch (err) {
        return fail(`docx_list_pending failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_approve ──────────────────────────────────────────────────────
  server.registerTool(
    "docx_approve",
    {
      description:
        "Approve a pending mutation. After approval the mutation is committed and shows up in the snapshot's history.",
      inputSchema: {
        handle: z.string(),
        mutation_id: z.string().describe("Mutation id from docx_list_pending."),
      },
    },
    async ({ handle, mutation_id }) => {
      try {
        const agent = lookupAgent(handle);
        agent.approveMutation(mutation_id);
        return ok({ approved: mutation_id, revision: agent.getSnapshot().revision });
      } catch (err) {
        return fail(`docx_approve failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_reject ───────────────────────────────────────────────────────
  server.registerTool(
    "docx_reject",
    {
      description:
        "Reject a pending mutation with an optional human-readable reason. The snapshot is unaffected; the mutation is dropped from the pending queue.",
      inputSchema: {
        handle: z.string(),
        mutation_id: z.string().describe("Mutation id from docx_list_pending."),
        reason: z.string().optional(),
      },
    },
    async ({ handle, mutation_id, reason }) => {
      try {
        const agent = lookupAgent(handle);
        agent.rejectMutation(mutation_id);
        return ok({ rejected: mutation_id, ...(reason ? { reason } : {}) });
      } catch (err) {
        return fail(`docx_reject failed: ${err instanceof Error ? err.message : String(err)}`);
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

  registerXlsxTools(server);

  // ── pptx_load ─────────────────────────────────────────────────────────
  server.registerTool(
    "pptx_load",
    {
      description:
        "Load a .pptx file from disk. Returns an opaque `handle` to use with subsequent pptx_* tools.",
      inputSchema: {
        path: z.string().describe("Absolute or workspace-relative path to a .pptx file."),
      },
    },
    async ({ path }) => {
      try {
        const buf = await readFile(resolve(path));
        const agent = await PptxAgent.fromBuffer(buf);
        const handle = randomUUID();
        pptxSessions.set(handle, agent);
        pptxSessionPaths.set(handle, resolve(path));
        return ok({ handle, path: resolve(path), summary: pptxInspectSnapshot(agent.getSnapshot()) });
      } catch (err) {
        return fail(`pptx_load failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pptx_save ─────────────────────────────────────────────────────────
  server.registerTool(
    "pptx_save",
    {
      description:
        "Serialize the current PPTX snapshot back to disk. Defaults to the path passed to pptx_load; pass `out_path` to write elsewhere.",
      inputSchema: {
        handle: z.string().describe("Handle returned by pptx_load."),
        out_path: z
          .string()
          .optional()
          .describe("Optional output path. Defaults to the original path passed to pptx_load."),
      },
    },
    async ({ handle, out_path }) => {
      try {
        const agent = lookupPptxAgent(handle);
        agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
        const target = out_path ? resolve(out_path) : pptxSessionPaths.get(handle);
        if (!target) {
          return fail(`pptx_save: no path known for handle "${handle}". Pass out_path explicitly.`);
        }
        const buf = Buffer.from(await agent.exportFile());
        await writeFile(target, buf);
        return ok({ wrote: target, bytes: buf.byteLength, revision: agent.getSnapshot().revision });
      } catch (err) {
        return fail(`pptx_save failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pptx_inspect ──────────────────────────────────────────────────────
  server.registerTool(
    "pptx_inspect",
    {
      description: "Return a structural summary (slide count, shape kinds, masters/layouts, parts).",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupPptxAgent(handle);
        return ok(pptxInspectSnapshot(agent.getSnapshot()));
      } catch (err) {
        return fail(`pptx_inspect failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pptx_get_text ─────────────────────────────────────────────────────
  server.registerTool(
    "pptx_get_text",
    {
      description:
        "Return the presentation as Markdown (default), structured JSON, or plain text. Optionally restrict to a single slide.",
      inputSchema: {
        handle: z.string(),
        format: z.enum(["markdown", "json", "text"]).optional().default("markdown"),
        slide: z.number().int().min(0).optional().describe("Restrict to a single 0-based slide index."),
      },
    },
    async ({ handle, format, slide }) => {
      try {
        const agent = lookupPptxAgent(handle);
        const snap = agent.getSnapshot();
        const range = slide !== undefined ? { startSlide: slide, endSlide: slide + 1 } : undefined;
        const fmt = format ?? "markdown";
        switch (fmt) {
          case "markdown": {
            // Markdown projection lives in @officeai/pptx; for slice, fall back to JSON.
            if (!range) return ok({ format: fmt, content: agent.toMarkdown() });
            const proj = pptxSnapshotToJsonProjection(snap, range);
            const lines = ["# Presentation"];
            for (const s of proj.slides) {
              lines.push(`## Slide ${s.index + 1} — \`${s.partPath}\` (slideId=${s.slideId})`);
              if (s.transition) {
                const speed = s.transition.speed ? ` (${s.transition.speed})` : "";
                lines.push(`- _transition_: **${s.transition.kind}**${speed}`);
              }
              if (s.animations && s.animations.length > 0) {
                lines.push(`- _animations_:`);
                for (const a of s.animations) {
                  const dur = a.durationMs !== undefined ? ` ${a.durationMs}ms` : "";
                  lines.push(
                    `  - \`${a.id}\` ${a.order + 1}. **${a.effect}**${dur} → cNvPr=${a.targetCNvPrId}`
                  );
                }
              }
              for (const sh of s.shapes) {
                if (sh.kind === "text" && sh.text) lines.push(`> ${sh.text.replaceAll("\n", " · ")}`);
                if (sh.kind === "table" && sh.table) {
                  for (const row of sh.table.cells) {
                    lines.push(
                      `| ${row.map((c) => (c.length > 0 ? c.replaceAll("\n", " · ") : "(empty)")).join(" | ")} |`
                    );
                  }
                }
                if (sh.kind === "chart" && sh.chart) {
                  lines.push(
                    `> chart (${sh.chart.chartType})${sh.chart.title ? ` — ${sh.chart.title}` : ""}`
                  );
                  if (sh.chart.categories.length > 0) {
                    lines.push(`> categories: ${sh.chart.categories.join(", ")}`);
                  }
                  for (const ser of sh.chart.series) {
                    lines.push(`> ${ser.name ? `${ser.name}: ` : ""}[${ser.values.join(", ")}]`);
                  }
                }
              }
            }
            return ok({ format: fmt, content: lines.join("\n") });
          }
          case "json":
            return ok(pptxSnapshotToJsonProjection(snap, range));
          case "text": {
            const proj = pptxSnapshotToJsonProjection(snap, range);
            const lines: string[] = [];
            for (const s of proj.slides) {
              for (const sh of s.shapes) {
                if (sh.kind === "text" && sh.text) lines.push(sh.text);
                if (sh.kind === "table" && sh.table) {
                  for (const row of sh.table.cells) {
                    for (const cell of row) {
                      if (cell.length > 0) lines.push(cell);
                    }
                  }
                }
                if (sh.kind === "chart" && sh.chart && sh.chart.title) {
                  lines.push(sh.chart.title);
                }
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
        return fail(`pptx_get_text failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pptx_search ───────────────────────────────────────────────────────
  server.registerTool(
    "pptx_search",
    {
      description: "Search every slide's text content for a query. Optional case-sensitive and regex flags.",
      inputSchema: {
        handle: z.string(),
        query: z.string().min(1),
        case_sensitive: z.boolean().optional().default(false),
        regex: z.boolean().optional().default(false),
      },
    },
    async ({ handle, query, case_sensitive, regex }) => {
      try {
        const agent = lookupPptxAgent(handle);
        const matches = agent.search({
          query,
          caseSensitive: case_sensitive ?? false,
          regex: regex ?? false,
        });
        return ok({ matches });
      } catch (err) {
        return fail(`pptx_search failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pptx_apply_command ────────────────────────────────────────────────
  server.registerTool(
    "pptx_apply_command",
    {
      description:
        "Apply a typed pptx command (e.g. `pptx:set-text`, `pptx:add-slide`, `pptx:insert-image`, `pptx:set-chart-title`, `pptx:set-slide-transition`, `pptx:add-shape-animation`). Pass an arbitrary payload object — schemas live in `@officeai/pptx/commands/payloads`.",
      inputSchema: {
        handle: z.string(),
        type: z.string().describe('Command type, e.g. "pptx:set-text"'),
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
        const agent = lookupPptxAgent(handle);
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
        return fail(`pptx_apply_command failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pptx_diff ─────────────────────────────────────────────────────────
  server.registerTool(
    "pptx_diff",
    {
      description:
        "Diff two loaded PPTX handles, OR diff a handle against the file on disk it was loaded from. Pass either {before, after} (two handles) or {handle, against: 'disk'}.",
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
          const a = lookupPptxAgent(before);
          const b = lookupPptxAgent(after);
          return ok(pptxDiffSnapshots(a.getSnapshot(), b.getSnapshot()));
        }
        if (handle && against === "disk") {
          const agent = lookupPptxAgent(handle);
          const path = pptxSessionPaths.get(handle);
          if (!path) return fail(`pptx_diff: no on-disk path for handle "${handle}".`);
          const buf = await readFile(path);
          const baseline = await PptxAgent.fromBuffer(buf);
          return ok(pptxDiffSnapshots(baseline.getSnapshot(), agent.getSnapshot()));
        }
        return fail("pptx_diff: pass either {before, after} (two handles) or {handle, against: 'disk'}.");
      } catch (err) {
        return fail(`pptx_diff failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  return server;
}

// ──────────────────────────────────────────────────────────────────────────
// xlsx_* tools
// ──────────────────────────────────────────────────────────────────────────

function registerXlsxTools(server: McpServer): void {
  // ── xlsx_load ─────────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_load",
    {
      description: "Load a .xlsx file from disk. Returns an opaque `handle` to use with subsequent tools.",
      inputSchema: {
        path: z.string().describe("Absolute or workspace-relative path to a .xlsx file."),
      },
    },
    async ({ path }) => {
      try {
        const buf = await readFile(resolve(path));
        const agent = await XlsxAgent.fromBuffer(buf);
        const handle = randomUUID();
        xlsxSessions.set(handle, agent);
        xlsxSessionPaths.set(handle, resolve(path));
        return ok({ handle, path: resolve(path), summary: inspectXlsxSnapshot(agent.getSnapshot()) });
      } catch (err) {
        return fail(`xlsx_load failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_save ─────────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_save",
    {
      description:
        "Serialize the current xlsx snapshot back to disk. Defaults to the path passed to xlsx_load; pass `out_path` to write elsewhere.",
      inputSchema: {
        handle: z.string().describe("Handle returned by xlsx_load."),
        out_path: z
          .string()
          .optional()
          .describe("Optional output path. Defaults to the original path passed to xlsx_load."),
      },
    },
    async ({ handle, out_path }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
        const target = out_path ? resolve(out_path) : xlsxSessionPaths.get(handle);
        if (!target) {
          return fail(`xlsx_save: no path known for handle "${handle}". Pass out_path explicitly.`);
        }
        const buf = Buffer.from(await agent.exportFile());
        await writeFile(target, buf);
        return ok({ wrote: target, bytes: buf.byteLength, revision: agent.getSnapshot().revision });
      } catch (err) {
        return fail(`xlsx_save failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_inspect ──────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_inspect",
    {
      description: "Return a structural summary (sheets, cells, parts, comments, merges).",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        return ok(inspectXlsxSnapshot(agent.getSnapshot()));
      } catch (err) {
        return fail(`xlsx_inspect failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_list_sheets ──────────────────────────────────────────────────
  server.registerTool(
    "xlsx_list_sheets",
    {
      description: "List sheets in tab order, with name, index, kind, and visibility state.",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        return ok({ sheets: agent.listSheets() });
      } catch (err) {
        return fail(`xlsx_list_sheets failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_get_text ─────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_get_text",
    {
      description: "Return one or more sheets as Markdown (default) or as a sparse JSON cell projection.",
      inputSchema: {
        handle: z.string(),
        format: z.enum(["markdown", "json"]).optional().default("markdown"),
        sheet: z.string().optional(),
        range: z.string().optional().describe("Optional A1 range; when set, also requires `sheet`."),
        max_rows: z.number().int().positive().optional(),
        max_cols: z.number().int().positive().optional(),
      },
    },
    async ({ handle, format, sheet, range, max_rows, max_cols }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        const fmt = format ?? "markdown";
        switch (fmt) {
          case "markdown": {
            const md = agent.toMarkdown({
              ...(sheet ? { sheet } : {}),
              ...(max_rows !== undefined ? { maxRows: max_rows } : {}),
              ...(max_cols !== undefined ? { maxCols: max_cols } : {}),
            });
            return ok({ format: fmt, content: md });
          }
          case "json":
            return ok(xlsxRangeToJson(agent, sheet, range));
          default: {
            const _exhaustive: never = fmt;
            void _exhaustive;
            return fail(`unknown format: ${String(format)}`);
          }
        }
      } catch (err) {
        return fail(`xlsx_get_text failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_search ───────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_search",
    {
      description: "Search workbook cells for text. Optional sheet filter, case-sensitive, and regex flags.",
      inputSchema: {
        handle: z.string(),
        query: z.string().min(1),
        sheet: z.string().optional(),
        case_sensitive: z.boolean().optional().default(false),
        regex: z.boolean().optional().default(false),
      },
    },
    async ({ handle, query, sheet, case_sensitive, regex }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        const results = agent.search({
          query,
          caseSensitive: case_sensitive ?? false,
          regex: regex ?? false,
          ...(sheet ? { sheet } : {}),
        });
        return ok({ matches: results });
      } catch (err) {
        return fail(`xlsx_search failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_apply_command ────────────────────────────────────────────────
  server.registerTool(
    "xlsx_apply_command",
    {
      description:
        "Apply an xlsx command (e.g. `xlsx:set-cell-value`, `xlsx:add-sheet`). Pass an arbitrary payload object — schemas live in `@officeai/xlsx/commands/payloads`.",
      inputSchema: {
        handle: z.string(),
        type: z.string().describe('Command type, e.g. "xlsx:set-cell-value"'),
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
        return await applyXlsxCommand(handle, type, payload, source, agent_id, auto_approve);
      } catch (err) {
        return fail(`xlsx_apply_command failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_list_pending ─────────────────────────────────────────────────
  server.registerTool(
    "xlsx_list_pending",
    {
      description:
        "List xlsx mutations still in `pending` state on the bus (typically agent-authored writes invoked with auto_approve=false).",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        const pending = agent.getPendingMutations().map((m) => ({
          id: m.id,
          command: { type: m.command.type, source: m.command.source },
          ...(m.command.source === "agent" && "agentId" in m.command ? { agentId: m.command.agentId } : {}),
          revision: m.after.revision,
          status: m.status,
        }));
        return ok({ pending });
      } catch (err) {
        return fail(`xlsx_list_pending failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_approve ──────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_approve",
    {
      description:
        "Approve a pending xlsx mutation. After approval the mutation is committed and shows up in the snapshot's history.",
      inputSchema: {
        handle: z.string(),
        mutation_id: z.string().describe("Mutation id from xlsx_list_pending."),
      },
    },
    async ({ handle, mutation_id }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        agent.approveMutation(mutation_id);
        return ok({ approved: mutation_id, revision: agent.getSnapshot().revision });
      } catch (err) {
        return fail(`xlsx_approve failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_reject ───────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_reject",
    {
      description:
        "Reject a pending xlsx mutation with an optional human-readable reason. The snapshot is unaffected; the mutation is dropped from the pending queue.",
      inputSchema: {
        handle: z.string(),
        mutation_id: z.string().describe("Mutation id from xlsx_list_pending."),
        reason: z.string().optional(),
      },
    },
    async ({ handle, mutation_id, reason }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        agent.rejectMutation(mutation_id);
        return ok({ rejected: mutation_id, ...(reason ? { reason } : {}) });
      } catch (err) {
        return fail(`xlsx_reject failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_undo ─────────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_undo",
    {
      description:
        "Undo the most recent approved mutation on this xlsx handle. No-op when the history is empty. Returns the resulting revision and a `did_undo` flag.",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        const m = agent.undo();
        return ok({
          did_undo: m !== null,
          revision: agent.getSnapshot().revision,
          can_undo: agent.canUndo(),
          can_redo: agent.canRedo(),
          undone: m ? { id: m.id, type: m.command.type } : null,
        });
      } catch (err) {
        return fail(`xlsx_undo failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_redo ─────────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_redo",
    {
      description:
        "Re-apply the most recently undone mutation on this xlsx handle. No-op when the redo stack is empty (e.g. immediately after a fresh authored mutation, which kills the redo trail).",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        const m = agent.redo();
        return ok({
          did_redo: m !== null,
          revision: agent.getSnapshot().revision,
          can_undo: agent.canUndo(),
          can_redo: agent.canRedo(),
          redone: m ? { id: m.id, type: m.command.type, status: m.status } : null,
        });
      } catch (err) {
        return fail(`xlsx_redo failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_diff ─────────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_diff",
    {
      description:
        "Diff two loaded xlsx handles, OR diff a handle against the file on disk it was loaded from. Pass either {before, after} (two handles) or {handle, against: 'disk'}.",
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
          const a = lookupXlsxAgent(before);
          const b = lookupXlsxAgent(after);
          return ok(diffXlsxSnapshots(a.getSnapshot(), b.getSnapshot()));
        }
        if (handle && against === "disk") {
          const agent = lookupXlsxAgent(handle);
          const path = xlsxSessionPaths.get(handle);
          if (!path) return fail(`xlsx_diff: no on-disk path for handle "${handle}".`);
          const buf = await readFile(path);
          const baseline = await XlsxAgent.fromBuffer(buf);
          return ok(diffXlsxSnapshots(baseline.getSnapshot(), agent.getSnapshot()));
        }
        return fail("xlsx_diff: pass either {before, after} (two handles) or {handle, against: 'disk'}.");
      } catch (err) {
        return fail(`xlsx_diff failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── Convenience write tools ───────────────────────────────────────────
  // Each one collapses internally to xlsx_apply_command but keeps the
  // wire payload trivial for LLM clients.

  const convenienceCommonSchema = {
    source: z.enum(["agent", "human", "system"]).optional().default("agent"),
    agent_id: z.string().optional().default("officeai-mcp"),
    auto_approve: z.boolean().optional().default(true),
  } as const;

  server.registerTool(
    "xlsx_set_cell",
    {
      description: "Set a single cell's literal value (collapses to xlsx:set-cell-value).",
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        ref: z.string().describe("A1 single-cell ref, e.g. 'B2'"),
        value: z.unknown(),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, sheet, ref, value, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:set-cell-value",
          { sheet, ref, value },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_set_cell failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "xlsx_set_formula",
    {
      description: "Set a single cell's formula (collapses to xlsx:set-cell-formula).",
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        ref: z.string(),
        formula: z.string(),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, sheet, ref, formula, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:set-cell-formula",
          { sheet, ref, formula },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_set_formula failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "xlsx_set_range",
    {
      description: "Set a 2-D matrix of cell values (collapses to xlsx:set-range-values).",
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        range: z.string(),
        values: z.array(z.array(z.unknown())),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, sheet, range, values, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:set-range-values",
          { sheet, range, values },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_set_range failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "xlsx_set_format",
    {
      description: "Apply a CellFormatPatch to a range (collapses to xlsx:set-cell-format).",
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        range: z.string(),
        format: z.record(z.string(), z.unknown()),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, sheet, range, format, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:set-cell-format",
          { sheet, range, format },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_set_format failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "xlsx_add_sheet",
    {
      description: "Append (or insert at `at`) a new worksheet (collapses to xlsx:add-sheet).",
      inputSchema: {
        handle: z.string(),
        name: z.string(),
        at: z.number().int().nonnegative().optional(),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, name, at, source, agent_id, auto_approve }) => {
      try {
        const payload: Record<string, unknown> = { name };
        if (at !== undefined) payload.at = at;
        return await applyXlsxCommand(handle, "xlsx:add-sheet", payload, source, agent_id, auto_approve);
      } catch (err) {
        return fail(`xlsx_add_sheet failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "xlsx_rename_sheet",
    {
      description: "Rename a worksheet (collapses to xlsx:rename-sheet).",
      inputSchema: {
        handle: z.string(),
        name: z.string(),
        new_name: z.string(),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, name, new_name, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:rename-sheet",
          { name, newName: new_name },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_rename_sheet failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  registerStructuralTool(server, "xlsx_insert_row", "xlsx:insert-row", "Insert N blank rows");
  registerStructuralTool(server, "xlsx_insert_column", "xlsx:insert-column", "Insert N blank columns");
  registerStructuralTool(server, "xlsx_delete_row", "xlsx:delete-row", "Delete N rows");
  registerStructuralTool(server, "xlsx_delete_column", "xlsx:delete-column", "Delete N columns");

  server.registerTool(
    "xlsx_merge",
    {
      description: "Merge an A1 range covering ≥2 cells (collapses to xlsx:merge-cells).",
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        range: z.string(),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, sheet, range, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:merge-cells",
          { sheet, range },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_merge failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "xlsx_unmerge",
    {
      description: "Unmerge an existing merged range (collapses to xlsx:unmerge-cells).",
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        range: z.string(),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, sheet, range, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:unmerge-cells",
          { sheet, range },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_unmerge failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "xlsx_add_comment",
    {
      description: "Attach a classic note to a single cell (collapses to xlsx:add-comment).",
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        ref: z.string(),
        text: z.string(),
        author: z.string(),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, sheet, ref, text, author, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:add-comment",
          { sheet, ref, text, author },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_add_comment failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );
}

function registerStructuralTool(
  server: McpServer,
  toolName: string,
  commandType: string,
  description: string
): void {
  server.registerTool(
    toolName,
    {
      description: `${description} starting at a 1-based index (collapses to ${commandType}).`,
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        at: z.number().int().positive(),
        count: z.number().int().positive(),
        source: z.enum(["agent", "human", "system"]).optional().default("agent"),
        agent_id: z.string().optional().default("officeai-mcp"),
        auto_approve: z.boolean().optional().default(true),
      },
    },
    async ({ handle, sheet, at, count, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          commandType,
          { sheet, at, count },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`${toolName} failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );
}

/**
 * Shared apply pipeline used by `xlsx_apply_command` and every
 * convenience tool. Returns an `ok(...)` payload describing the
 * mutation status (with the post-auto-approve resolution) and the new
 * snapshot revision.
 */
async function applyXlsxCommand(
  handle: string,
  type: string,
  payload: unknown,
  source: "agent" | "human" | "system" | undefined,
  agentId: string | undefined,
  autoApprove: boolean | undefined
): Promise<ReturnType<typeof ok>> {
  const agent = lookupXlsxAgent(handle);
  const effectiveSource = source ?? "agent";
  const mutation = await agent.applyCommand({
    type,
    payload,
    source: effectiveSource,
    ...(effectiveSource === "agent" ? { agentId: agentId ?? "officeai-mcp" } : {}),
  });
  if ((autoApprove ?? true) && mutation.status === "pending") {
    agent.approveMutation(mutation.id);
  }
  return ok({
    mutation: {
      id: mutation.id,
      status: agent.getPendingMutations().some((m) => m.id === mutation.id) ? "pending" : mutation.status,
      ...(mutation.rejection ? { rejection: mutation.rejection } : {}),
    },
    revision: agent.getSnapshot().revision,
  });
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
