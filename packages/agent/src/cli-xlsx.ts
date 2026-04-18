/**
 * `office-agent xlsx …` subcommand group.
 *
 * Mirrors the docx surface in `cli.ts`: a per-format inspect/read/search
 * trio, a typed write subcommand for every `xlsx:*` payload kind, plus
 * a generic `xlsx apply` escape hatch and an on-disk `xlsx diff`.
 *
 * The MCP server in `mcp.ts` re-uses the helpers exported from this
 * module (`loadXlsxAgent`, `inspectXlsxSnapshot`, `xlsxRangeToJson`,
 * `diffXlsxOnDisk`) so the two surfaces report identical structures.
 */

import { readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { Command, Option } from "commander";
import { XlsxAgent, diffXlsxSnapshots, type XlsxRangeSnapshot, type XlsxSnapshot } from "@officeai/xlsx";
import type { DocumentDiff, Mutation } from "@officeai/core";
import { CliError, parseIntOpt, stringifyJson, type IO } from "./cli-shared.js";

// ──────────────────────────────────────────────────────────────────────────
// Agent loading + structural projections
// ──────────────────────────────────────────────────────────────────────────

export async function loadXlsxAgent(input: string): Promise<XlsxAgent> {
  const buf = await readFile(resolve(input));
  return XlsxAgent.fromBuffer(buf);
}

export interface XlsxSheetSummary {
  readonly name: string;
  readonly index: number;
  readonly kind: "worksheet" | "non-worksheet";
  readonly state: "visible" | "hidden" | "veryHidden";
  readonly cellCount: number;
  readonly mergeCount: number;
  readonly commentCount: number;
}

export interface XlsxSnapshotSummary {
  readonly format: "xlsx";
  readonly revision: number;
  readonly sheets: ReadonlyArray<XlsxSheetSummary>;
  readonly partCount: number;
  readonly opaquePartCount: number;
  readonly comments: number;
  readonly merges: number;
  readonly date1904: boolean;
  readonly parts: ReadonlyArray<string>;
}

export function inspectXlsxSnapshot(snap: XlsxSnapshot): XlsxSnapshotSummary {
  let totalComments = 0;
  let totalMerges = 0;
  const sheets: XlsxSheetSummary[] = snap.root.sheets.map((s) => {
    totalComments += s.comments.length;
    totalMerges += s.merges.length;
    return {
      name: s.name,
      index: s.index,
      kind: s.kind,
      state: s.state,
      cellCount: s.cells.size,
      mergeCount: s.merges.length,
      commentCount: s.comments.length,
    };
  });
  return {
    format: "xlsx",
    revision: snap.revision,
    sheets,
    partCount: snap.container.parts.size,
    opaquePartCount: snap.root.opaqueParts.size,
    comments: totalComments,
    merges: totalMerges,
    date1904: snap.root.date1904,
    parts: Array.from(snap.container.parts.keys()).sort(),
  };
}

/**
 * JSON projection of a sheet/range. When `range` is omitted the
 * sheet's bounding box is returned (consistent with `agent.getRange`'s
 * `xlsx-sheet` shape).
 */
export function xlsxRangeToJson(
  agent: XlsxAgent,
  sheet?: string,
  range?: string
): {
  format: "xlsx";
  revision: number;
  sheet: string;
  range: string;
  rows: number;
  cols: number;
  cells: XlsxRangeSnapshot["cells"];
} {
  const snap = agent.getSnapshot();
  const targetSheet = sheet ?? firstWorksheetName(snap);
  if (!targetSheet) {
    throw new CliError(2, "xlsx: workbook has no worksheets to project");
  }
  const projection = range
    ? agent.getRange({ kind: "xlsx-range", sheet: targetSheet, range })
    : agent.getRange({ kind: "xlsx-sheet", sheet: targetSheet });
  return {
    format: "xlsx",
    revision: snap.revision,
    sheet: projection.sheet,
    range: projection.range,
    rows: projection.rows,
    cols: projection.cols,
    cells: projection.cells,
  };
}

export function xlsxRangeToMarkdown(
  agent: XlsxAgent,
  sheet?: string,
  range?: string,
  maxRows?: number,
  maxCols?: number
): string {
  if (range) {
    if (!sheet) throw new CliError(64, "--range requires --sheet");
    const snap = agent.getRange({ kind: "xlsx-range", sheet, range });
    return rangeSnapshotToMarkdown(snap);
  }
  return agent.toMarkdown({
    ...(sheet ? { sheet } : {}),
    ...(maxRows !== undefined ? { maxRows } : {}),
    ...(maxCols !== undefined ? { maxCols } : {}),
  });
}

/** Diff two on-disk xlsx files, parsing both via `XlsxAgent.fromBuffer`. */
export async function diffXlsxOnDisk(before: string, after: string): Promise<DocumentDiff> {
  const beforeAgent = await loadXlsxAgent(before);
  const afterAgent = await loadXlsxAgent(after);
  return diffXlsxSnapshots(beforeAgent.getSnapshot(), afterAgent.getSnapshot());
}

function firstWorksheetName(snap: XlsxSnapshot): string | undefined {
  return snap.root.sheets.find((s) => s.kind === "worksheet")?.name;
}

function rangeSnapshotToMarkdown(snap: XlsxRangeSnapshot): string {
  if (snap.cells.length === 0) return `## ${snap.sheet} ${snap.range}\n\n_(empty)_`;
  // Build a dense matrix for the projected window.
  const rowOffsets = snap.cells.map((c) => c.row);
  const colOffsets = snap.cells.map((c) => c.col);
  const minRow = Math.min(...rowOffsets);
  const minCol = Math.min(...colOffsets);
  const grid: string[][] = Array.from({ length: snap.rows }, () =>
    Array.from({ length: snap.cols }, () => "")
  );
  for (const c of snap.cells) {
    const text = cellValueToString(c.value);
    grid[c.row - minRow][c.col - minCol] = c.formula ? `=${c.formula}` : text;
  }
  const header: string[] = [""];
  for (let i = 0; i < snap.cols; i++) header.push(columnLetter(minCol + i));
  const lines: string[] = [
    `## ${snap.sheet} ${snap.range}`,
    "",
    `| ${header.join(" | ")} |`,
    `| ${header.map(() => "---").join(" | ")} |`,
  ];
  for (let r = 0; r < snap.rows; r++) {
    const row: string[] = [String(minRow + r + 1)];
    for (let c = 0; c < snap.cols; c++) row.push(grid[r][c]);
    lines.push(`| ${row.join(" | ")} |`);
  }
  return lines.join("\n");
}

function cellValueToString(v: XlsxRangeSnapshot["cells"][number]["value"]): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v;
  if (typeof v === "object" && "kind" in v) return v.code;
  return String(v);
}

function columnLetter(col: number): string {
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ──────────────────────────────────────────────────────────────────────────
// Shared write pipeline
// ──────────────────────────────────────────────────────────────────────────

interface XlsxWriteOpts {
  file: string;
  out?: string;
  source?: "agent" | "human";
  agentId?: string;
  approve?: boolean;
  pretty?: boolean;
}

/**
 * Shared write pipeline for every xlsx write subcommand. Loads the
 * file, dispatches `type`+`payload`, optionally auto-approves, then
 * serializes back to disk and prints a JSON summary. Mirrors the
 * `runWrite` helper in cli.ts so the two surfaces stay in sync.
 */
async function runXlsxWrite(io: IO, opts: XlsxWriteOpts, type: string, payload: unknown): Promise<void> {
  const agent = await loadXlsxAgent(opts.file);
  const source = opts.source ?? "agent";
  const m = await agent.applyCommand({
    type,
    payload,
    source,
    ...(source === "agent" ? { agentId: opts.agentId ?? "office-agent-cli" } : {}),
  });
  const approve = opts.approve !== false;
  if (approve) {
    agent.getPendingMutations().forEach((p) => agent.approveMutation(p.id));
  }
  const out = opts.out ?? opts.file;
  await writeFile(resolve(out), Buffer.from(await agent.exportFile()));
  io.stdout.write(
    stringifyJson(
      mutationSummary(
        approve
          ? m
          : ({
              ...(m as object),
              status: "pending",
            } as unknown as Mutation<XlsxSnapshot>),
        out
      ),
      opts.pretty === true
    ) + "\n"
  );
}

function mutationSummary(
  m: { id: string; status: string; rejection?: { code: string; message: string } | undefined },
  wrote: string
): {
  wrote: string;
  mutation: { id: string; status: string; rejection?: { code: string; message: string } };
} {
  return {
    wrote,
    mutation: {
      id: m.id,
      status: m.status,
      ...(m.rejection ? { rejection: m.rejection } : {}),
    },
  };
}

/**
 * Parse a JSON string from a CLI flag and surface readable errors.
 * Used by `xlsx set-cell --value <json>` and `xlsx apply --payload <json>`.
 */
function parseJsonFlag(raw: string, flag: string): unknown {
  try {
    return JSON.parse(raw);
  } catch (err) {
    throw new CliError(64, `${flag}: invalid JSON: ${(err as Error).message}`);
  }
}

function normalizeCommands(
  data: unknown
): Array<{ type: string; payload: unknown; source?: "human" | "agent" | "system"; agentId?: string }> {
  const list = Array.isArray(data)
    ? data
    : isObj(data) && Array.isArray((data as { commands?: unknown[] }).commands)
      ? (data as { commands: unknown[] }).commands
      : [data];
  return list.map((c) => {
    if (!isObj(c) || typeof c.type !== "string") {
      throw new Error("each command must be an object with a string `type`");
    }
    return {
      type: c.type as string,
      payload: c.payload,
      ...(c.source ? { source: c.source as "human" | "agent" | "system" } : { source: "agent" as const }),
      ...(typeof c.agentId === "string" ? { agentId: c.agentId } : { agentId: "office-agent-cli" }),
    };
  });
}

function isObj(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null;
}

// ──────────────────────────────────────────────────────────────────────────
// Subcommand registration
// ──────────────────────────────────────────────────────────────────────────

interface CommonWriteFlags {
  file: string;
  out?: string;
  source: "agent" | "human";
  agentId: string;
  approve: boolean;
  pretty: boolean;
}

export function registerXlsxSubcommands(xlsx: Command, io: IO): void {
  xlsx
    .command("inspect")
    .description("Print a structural summary (sheets, cells, parts, comments, merges) as JSON.")
    .requiredOption("--file <path>", "Path to a .xlsx file")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; pretty: boolean }) => {
      const agent = await loadXlsxAgent(opts.file);
      io.stdout.write(stringifyJson(inspectXlsxSnapshot(agent.getSnapshot()), opts.pretty) + "\n");
    });

  xlsx
    .command("read")
    .description("Render a sheet (or range) as Markdown or JSON.")
    .requiredOption("--file <path>", "Path to a .xlsx file")
    .option("--sheet <name>", "Sheet name (defaults to the first worksheet)")
    .option("--range <A1:B2>", "Optional A1 range; requires --sheet")
    .addOption(
      new Option("--format <fmt>", "Output format").choices(["markdown", "json"]).default("markdown")
    )
    .option("--pretty", "Pretty-print JSON output (only with --format json)", false)
    .option("--max-rows <n>", "Markdown: cap rows rendered per sheet", parseIntOpt)
    .option("--max-cols <n>", "Markdown: cap columns rendered per sheet", parseIntOpt)
    .action(
      async (opts: {
        file: string;
        sheet?: string;
        range?: string;
        format: "markdown" | "json";
        pretty: boolean;
        maxRows?: number;
        maxCols?: number;
      }) => {
        const agent = await loadXlsxAgent(opts.file);
        switch (opts.format) {
          case "markdown":
            io.stdout.write(
              xlsxRangeToMarkdown(agent, opts.sheet, opts.range, opts.maxRows, opts.maxCols) + "\n"
            );
            return;
          case "json":
            io.stdout.write(
              stringifyJson(xlsxRangeToJson(agent, opts.sheet, opts.range), opts.pretty) + "\n"
            );
            return;
          default: {
            const _exhaustive: never = opts.format;
            void _exhaustive;
          }
        }
      }
    );

  xlsx
    .command("search")
    .description("Search workbook cells for text and print matches as JSON.")
    .requiredOption("--file <path>", "Path to a .xlsx file")
    .requiredOption("-q, --query <text>", "Search query")
    .option("--sheet <name>", "Restrict to a single sheet")
    .option("--regex", "Treat the query as a regular expression", false)
    .option("--case-sensitive", "Case-sensitive search", false)
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        query: string;
        sheet?: string;
        regex: boolean;
        caseSensitive: boolean;
        pretty: boolean;
      }) => {
        const agent = await loadXlsxAgent(opts.file);
        const results = agent.search({
          query: opts.query,
          regex: opts.regex,
          caseSensitive: opts.caseSensitive,
          ...(opts.sheet ? { sheet: opts.sheet } : {}),
        });
        io.stdout.write(stringifyJson(results, opts.pretty) + "\n");
      }
    );

  // ── Typed write subcommands ───────────────────────────────────────────
  attachCommonWriteFlags(
    xlsx
      .command("set-cell")
      .description("Set a single cell's literal value.")
      .requiredOption("--file <path>", "Path to a .xlsx file")
      .requiredOption("--sheet <name>", "Sheet name")
      .requiredOption("--ref <A1>", "A1 single-cell ref, e.g. B7")
      .requiredOption(
        "--value <json>",
        'JSON-encoded literal (null, true, 42, "text", {"kind":"error","code":"#REF!"})'
      )
  ).action(async (opts: CommonWriteFlags & { sheet: string; ref: string; value: string }) => {
    const value = parseJsonFlag(opts.value, "--value");
    await runXlsxWrite(io, opts, "xlsx:set-cell-value", { sheet: opts.sheet, ref: opts.ref, value });
  });

  attachCommonWriteFlags(
    xlsx
      .command("set-formula")
      .description("Set a single cell's formula (with or without leading '=').")
      .requiredOption("--file <path>", "Path to a .xlsx file")
      .requiredOption("--sheet <name>", "Sheet name")
      .requiredOption("--ref <A1>", "A1 single-cell ref")
      .requiredOption("--formula <text>", "Formula text")
  ).action(async (opts: CommonWriteFlags & { sheet: string; ref: string; formula: string }) => {
    await runXlsxWrite(io, opts, "xlsx:set-cell-formula", {
      sheet: opts.sheet,
      ref: opts.ref,
      formula: opts.formula,
    });
  });

  attachCommonWriteFlags(
    xlsx
      .command("set-range")
      .description("Set a row-major 2-D matrix of cell values across a range.")
      .requiredOption("--file <path>", "Path to a .xlsx file")
      .requiredOption("--sheet <name>", "Sheet name")
      .requiredOption("--range <A1:C3>", "A1 range")
      .requiredOption("--values-json <json>", "JSON 2-D array; dims must equal --range")
  ).action(async (opts: CommonWriteFlags & { sheet: string; range: string; valuesJson: string }) => {
    const values = parseJsonFlag(opts.valuesJson, "--values-json");
    await runXlsxWrite(io, opts, "xlsx:set-range-values", {
      sheet: opts.sheet,
      range: opts.range,
      values,
    });
  });

  attachCommonWriteFlags(
    xlsx
      .command("set-format")
      .description("Apply a format patch to a range (font/fill/border/alignment/numberFormat).")
      .requiredOption("--file <path>", "Path to a .xlsx file")
      .requiredOption("--sheet <name>", "Sheet name")
      .requiredOption("--range <A1[:B2]>", "A1 single cell or range")
      .requiredOption("--format-json <json>", "JSON format patch (see CellFormatPatch)")
  ).action(async (opts: CommonWriteFlags & { sheet: string; range: string; formatJson: string }) => {
    const format = parseJsonFlag(opts.formatJson, "--format-json");
    await runXlsxWrite(io, opts, "xlsx:set-cell-format", {
      sheet: opts.sheet,
      range: opts.range,
      format,
    });
  });

  attachCommonWriteFlags(
    xlsx
      .command("add-sheet")
      .description("Append (or insert at --at) a new worksheet.")
      .requiredOption("--file <path>", "Path to a .xlsx file")
      .requiredOption("--name <name>", "New sheet name (Excel rules apply)")
      .option("--at <index>", "0-based insert position (defaults to append)", parseIntOpt)
  ).action(async (opts: CommonWriteFlags & { name: string; at?: number }) => {
    await runXlsxWrite(io, opts, "xlsx:add-sheet", {
      name: opts.name,
      ...(opts.at !== undefined ? { at: opts.at } : {}),
    });
  });

  attachCommonWriteFlags(
    xlsx
      .command("rename-sheet")
      .description("Rename a sheet by current name.")
      .requiredOption("--file <path>", "Path to a .xlsx file")
      .requiredOption("--name <current>", "Current sheet name")
      .requiredOption("--new-name <new>", "New sheet name")
  ).action(async (opts: CommonWriteFlags & { name: string; newName: string }) => {
    await runXlsxWrite(io, opts, "xlsx:rename-sheet", {
      name: opts.name,
      newName: opts.newName,
    });
  });

  attachCommonWriteFlags(
    xlsx
      .command("insert-row")
      .description("Insert N blank rows BEFORE a 1-based row index.")
      .requiredOption("--file <path>", "Path to a .xlsx file")
      .requiredOption("--sheet <name>", "Sheet name")
      .requiredOption("--at <n>", "1-based row index (insertion is BEFORE this row)", parseIntOpt)
      .requiredOption("--count <n>", "Number of blank rows to insert", parseIntOpt)
  ).action(async (opts: CommonWriteFlags & { sheet: string; at: number; count: number }) => {
    await runXlsxWrite(io, opts, "xlsx:insert-row", {
      sheet: opts.sheet,
      at: opts.at,
      count: opts.count,
    });
  });

  attachCommonWriteFlags(
    xlsx
      .command("insert-column")
      .description("Insert N blank columns BEFORE a 1-based column index.")
      .requiredOption("--file <path>", "Path to a .xlsx file")
      .requiredOption("--sheet <name>", "Sheet name")
      .requiredOption("--at <n>", "1-based column index", parseIntOpt)
      .requiredOption("--count <n>", "Number of blank columns to insert", parseIntOpt)
  ).action(async (opts: CommonWriteFlags & { sheet: string; at: number; count: number }) => {
    await runXlsxWrite(io, opts, "xlsx:insert-column", {
      sheet: opts.sheet,
      at: opts.at,
      count: opts.count,
    });
  });

  attachCommonWriteFlags(
    xlsx
      .command("delete-row")
      .description("Delete N rows starting at a 1-based row index.")
      .requiredOption("--file <path>", "Path to a .xlsx file")
      .requiredOption("--sheet <name>", "Sheet name")
      .requiredOption("--at <n>", "1-based row index of the first row to drop", parseIntOpt)
      .requiredOption("--count <n>", "Number of rows to drop", parseIntOpt)
  ).action(async (opts: CommonWriteFlags & { sheet: string; at: number; count: number }) => {
    await runXlsxWrite(io, opts, "xlsx:delete-row", {
      sheet: opts.sheet,
      at: opts.at,
      count: opts.count,
    });
  });

  attachCommonWriteFlags(
    xlsx
      .command("delete-column")
      .description("Delete N columns starting at a 1-based column index.")
      .requiredOption("--file <path>", "Path to a .xlsx file")
      .requiredOption("--sheet <name>", "Sheet name")
      .requiredOption("--at <n>", "1-based column index of the first column to drop", parseIntOpt)
      .requiredOption("--count <n>", "Number of columns to drop", parseIntOpt)
  ).action(async (opts: CommonWriteFlags & { sheet: string; at: number; count: number }) => {
    await runXlsxWrite(io, opts, "xlsx:delete-column", {
      sheet: opts.sheet,
      at: opts.at,
      count: opts.count,
    });
  });

  attachCommonWriteFlags(
    xlsx
      .command("merge-cells")
      .description("Merge an A1 range covering ≥2 cells.")
      .requiredOption("--file <path>", "Path to a .xlsx file")
      .requiredOption("--sheet <name>", "Sheet name")
      .requiredOption("--range <A1:B2>", "A1 range")
  ).action(async (opts: CommonWriteFlags & { sheet: string; range: string }) => {
    await runXlsxWrite(io, opts, "xlsx:merge-cells", { sheet: opts.sheet, range: opts.range });
  });

  attachCommonWriteFlags(
    xlsx
      .command("unmerge-cells")
      .description("Unmerge an existing merged range (must match exactly).")
      .requiredOption("--file <path>", "Path to a .xlsx file")
      .requiredOption("--sheet <name>", "Sheet name")
      .requiredOption("--range <A1:B2>", "Existing merge range")
  ).action(async (opts: CommonWriteFlags & { sheet: string; range: string }) => {
    await runXlsxWrite(io, opts, "xlsx:unmerge-cells", { sheet: opts.sheet, range: opts.range });
  });

  attachCommonWriteFlags(
    xlsx
      .command("add-comment")
      .description("Attach a classic note (single-cell anchor).")
      .requiredOption("--file <path>", "Path to a .xlsx file")
      .requiredOption("--sheet <name>", "Sheet name")
      .requiredOption("--ref <A1>", "Single-cell A1 ref (range refs are rejected)")
      .requiredOption("--text <text>", "Comment body (plain text)")
      .requiredOption("--author <name>", "Author")
  ).action(async (opts: CommonWriteFlags & { sheet: string; ref: string; text: string; author: string }) => {
    await runXlsxWrite(io, opts, "xlsx:add-comment", {
      sheet: opts.sheet,
      ref: opts.ref,
      text: opts.text,
      author: opts.author,
    });
  });

  // ── Generic escape hatch ──────────────────────────────────────────────
  attachCommonWriteFlags(
    xlsx
      .command("apply")
      .description("Generic command escape hatch — pass --type and --payload (JSON).")
      .requiredOption("--file <path>", "Path to a .xlsx file")
      .requiredOption("--type <command-type>", 'Command type, e.g. "xlsx:set-cell-value"')
      .requiredOption("--payload <json>", "JSON payload object")
  ).action(async (opts: CommonWriteFlags & { type: string; payload: string }) => {
    const payload = parseJsonFlag(opts.payload, "--payload");
    await runXlsxWrite(io, opts, opts.type, payload);
  });

  // ── Apply commands JSON file (matches `docx apply -c <path>`) ─────────
  xlsx
    .command("apply-file")
    .description("Apply a JSON command file (single command or { commands: [...] }) and write the result.")
    .requiredOption("--file <path>", "Path to a .xlsx file")
    .requiredOption("-c, --commands <path>", "Path to a JSON file containing one or more commands")
    .option("--out <path>", "Path to write the resulting .xlsx file (defaults to --file, in place)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; commands: string; out?: string; pretty: boolean }) => {
      const agent = await loadXlsxAgent(opts.file);
      const raw = await readFile(resolve(opts.commands), "utf8");
      const data: unknown = JSON.parse(raw);
      const cmds = normalizeCommands(data);
      const muts = await agent.applyCommands(cmds);
      agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
      const out = opts.out ?? opts.file;
      await writeFile(resolve(out), Buffer.from(await agent.exportFile()));
      io.stdout.write(
        stringifyJson(
          {
            wrote: out,
            mutations: muts.map((m) => ({ id: m.id, type: m.command.type, status: m.status })),
          },
          opts.pretty
        ) + "\n"
      );
    });

  xlsx
    .command("diff")
    .description("Compute a structural diff between two on-disk .xlsx files.")
    .requiredOption("--before <path>", "Path to the baseline .xlsx file")
    .requiredOption("--after <path>", "Path to the modified .xlsx file")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { before: string; after: string; pretty: boolean }) => {
      const diff = await diffXlsxOnDisk(opts.before, opts.after);
      io.stdout.write(stringifyJson(diff, opts.pretty) + "\n");
    });
}

/**
 * Glue for the "common write flags" applied to every typed write
 * subcommand. Centralizing them keeps the option list trivially in
 * sync between docx and xlsx surfaces.
 */
function attachCommonWriteFlags(cmd: Command): Command {
  return cmd
    .option("--out <path>", "Path to write the resulting .xlsx file (defaults to --file, in place)")
    .addOption(new Option("--source <src>", "Mutation source").choices(["agent", "human"]).default("agent"))
    .option("--agent-id <id>", "Agent identifier (defaults to office-agent-cli)", "office-agent-cli")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false);
}
