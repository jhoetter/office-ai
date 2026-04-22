/**
 * Action-catalogue → MCP tool adapter.
 *
 * Counterpart to `actions-to-cli.ts`. Walks an `ActionDescriptor[]`
 * and registers an MCP tool for every entry that:
 *   • lists `"cli"` in `surfaces` (the CLI surface is our proxy for
 *     "agent-callable" — anything reachable from the CLI is by
 *     definition reachable headlessly), AND
 *   • declares a non-null `commandType`, AND
 *   • declares its `args` and `buildPayload` (so we can synthesise a
 *     zod input schema and the bus payload mechanically).
 *
 * Each generated tool is named `<format>.<id-without-prefix>` (mirroring
 * the CLI subcommand naming scheme), takes an opaque agent `handle`
 * plus the catalogue's declared args, dispatches the bus command,
 * optionally writes the result, and returns a JSON summary.
 *
 * Hand-rolled MCP tools for non-mutating reads (`docx_get_text`,
 * `xlsx_get_range_json`, …) live in `mcp.ts` and stay authoritative —
 * the auto-binder NEVER overlaps with them because read-only catalogue
 * entries have `commandType: null`.
 *
 * Entries that need bespoke schemas (heavily structured payloads like
 * the FillSpec for `pptx:set-slide-background`) opt out simply by
 * leaving `args` / `buildPayload` blank.
 */

import { z, type ZodTypeAny } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { ActionArg, ActionDescriptor } from "@officeai/core";

/**
 * The minimal contract the host (mcp.ts) provides so the adapter can
 * resolve a handle, dispatch a command, and (if requested) write the
 * result without dragging the per-format agent type into this file.
 */
export interface McpDispatchContext {
  /** Format key, used in tool names and error messages. */
  readonly format: "docx" | "xlsx" | "pptx";
  /**
   * Look up an opaque agent handle. The returned value is forwarded to
   * `dispatch` unchanged.
   */
  lookup(handle: string): unknown;
  /**
   * Dispatch one bus command and (optionally) write the resulting
   * snapshot back to disk. Returns a JSON-serialisable summary.
   */
  dispatch(args: {
    readonly handle: string;
    readonly agent: unknown;
    readonly commandType: string;
    readonly payload: unknown;
    readonly outPath?: string;
    readonly source: "agent" | "human";
    readonly approve: boolean;
  }): Promise<unknown>;
}

/**
 * Register every catalogue entry that opted into the `cli` surface
 * AND has the metadata needed for auto-generation. Returns the list
 * of action ids that were registered (for diagnostics / parity).
 */
export function registerActionsAsMcpTools(
  server: McpServer,
  actions: ReadonlyArray<ActionDescriptor>,
  ctx: McpDispatchContext
): { registered: ReadonlyArray<string>; skipped: ReadonlyArray<string> } {
  const registered: string[] = [];
  const skipped: string[] = [];

  for (const action of actions) {
    if (!action.surfaces.includes("cli")) {
      skipped.push(action.id);
      continue;
    }
    if (action.hidden) {
      skipped.push(action.id);
      continue;
    }
    if (action.commandType === null || action.commandType === undefined) {
      // Non-mutating reads are exposed by hand-rolled tools in mcp.ts.
      skipped.push(action.id);
      continue;
    }
    if (!action.args || !action.buildPayload) {
      skipped.push(action.id);
      continue;
    }
    try {
      registerOne(server, action, ctx);
      registered.push(action.id);
    } catch (err) {
      // Hand-rolled MCP tool blocks in mcp.ts win when both attempt
      // the same name. The MCP SDK throws "Tool X is already
      // registered" — swallow it so adding catalogue args/buildPayload
      // to an entry that has a hand-rolled wrapper doesn't crash
      // server bootstrap.
      const message = err instanceof Error ? err.message : String(err);
      if (!/already registered/i.test(message)) {
        throw err;
      }
      skipped.push(action.id);
    }
  }

  return { registered, skipped };
}

function registerOne(server: McpServer, action: ActionDescriptor, ctx: McpDispatchContext): void {
  const toolName = mcpToolName(action.id);
  const inputSchema = buildInputSchema(action);
  const description = buildDescription(action);

  server.registerTool(
    toolName,
    { description, inputSchema },
    async (input: Record<string, unknown>) => {
      try {
        const handle = String(input.handle ?? "");
        if (!handle) throw new Error(`${toolName}: missing required "handle" argument`);
        const agent = ctx.lookup(handle);
        const buildPayload = action.buildPayload;
        if (!buildPayload) {
          throw new Error(`${toolName}: catalogue entry has no buildPayload (internal error)`);
        }
        const declared = action.args ?? [];
        const parsed = collectArgs(input, declared);
        const payload = buildPayload(parsed);
        const result = await ctx.dispatch({
          handle,
          agent,
          commandType: action.commandType as string,
          payload,
          outPath: typeof input.out_path === "string" ? input.out_path : undefined,
          source: input.source === "human" ? "human" : "agent",
          approve: input.approve !== false,
        });
        const text = JSON.stringify(result, null, 2);
        return {
          content: [{ type: "text" as const, text }],
          structuredContent: result as Record<string, unknown>,
        };
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        return {
          isError: true as const,
          content: [{ type: "text" as const, text: `${toolName} failed: ${msg}` }],
        };
      }
    }
  );
}

/**
 * `docx.delete-row` → `docx_delete_row`. MCP convention is
 * snake-case-ish for tool names, and several existing hand-rolled
 * tools (`docx_load`, `xlsx_get_range_json`) already follow the
 * `<format>_<verb>_<noun>` convention.
 */
export function mcpToolName(actionId: string): string {
  const local = actionId.includes(".") ? actionId.slice(actionId.indexOf(".") + 1) : actionId;
  const format = actionId.includes(".") ? actionId.slice(0, actionId.indexOf(".")) : "";
  const slug = local.replace(/-/g, "_").replace(/\./g, "_");
  return format ? `${format}_${slug}` : slug;
}

function buildDescription(action: ActionDescriptor): string {
  const lines: string[] = [];
  lines.push(action.description);
  lines.push("");
  lines.push(`Handle is required (call ${action.id.split(".")[0]}_load first).`);
  lines.push(
    `If you want the result persisted, pass \`out_path\` or call ${action.id.split(".")[0]}_save afterwards.`
  );
  return lines.join("\n");
}

/**
 * Build a zod object schema for the tool's inputs. Always includes the
 * common `handle`, `out_path`, `source`, `approve` fields the dispatch
 * context expects, plus one zod field per declared arg.
 */
function buildInputSchema(action: ActionDescriptor): Record<string, ZodTypeAny> {
  const out: Record<string, ZodTypeAny> = {
    handle: z.string().describe("Opaque agent handle returned by the matching <format>_load tool."),
    out_path: z.string().optional().describe("Optional output path. If unset, no write happens here."),
    source: z.enum(["agent", "human"]).optional().describe("Mutation source (default: agent)."),
    approve: z
      .boolean()
      .optional()
      .describe("If false, leave the resulting mutation pending instead of auto-approving (default: true)."),
  };
  for (const arg of action.args ?? []) {
    out[arg.name] = zodForArg(arg);
  }
  return out;
}

function zodForArg(arg: ActionArg): ZodTypeAny {
  let base: ZodTypeAny;
  switch (arg.kind) {
    case "string":
    case "selector":
    case "filepath":
      base = z.string();
      break;
    case "enum":
      base = arg.choices && arg.choices.length > 0
        ? z.enum(arg.choices as [string, ...string[]])
        : z.string();
      break;
    case "number":
      base = z.number();
      break;
    case "boolean":
      base = z.boolean();
      break;
    case "stringList":
      base = z.array(z.string());
      break;
    default: {
      const _exhaustive: never = arg.kind;
      void _exhaustive;
      base = z.unknown();
    }
  }
  if (arg.description) base = base.describe(arg.description);
  if (!arg.required) base = base.optional();
  if (arg.default !== undefined) {
    // Note: zod's `.default()` only runs when the field is `undefined`.
    // We don't apply it for booleans where the default is `false` (the
    // adapter treats absence as false anyway).
    if (typeof arg.default === "string" || typeof arg.default === "number") {
      base = (base as z.ZodOptional<ZodTypeAny>).default(arg.default);
    }
  }
  return base;
}

function collectArgs(opts: Record<string, unknown>, args: ReadonlyArray<ActionArg>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const a of args) {
    out[a.name] = opts[a.name];
  }
  return out;
}
