#!/usr/bin/env node
import { readFile, writeFile } from "node:fs/promises";
import { extname, resolve } from "node:path";
import { Command, Option } from "commander";
import { DocxAgent, paragraphPlainText } from "@officeai/docx";
import type { DocxComment, DocxPosition, DocxSnapshot, Paragraph } from "@officeai/docx";
import type { CommandLite, Mutation } from "@officeai/core";
import { parseSelector, SelectorError, type Selector } from "./selector.js";
import { runMcpStdioServer } from "./mcp.js";
import { CliError, parseIntOpt, stringifyJson, type IO } from "./cli-shared.js";
import { registerXlsxSubcommands } from "./cli-xlsx.js";
import { registerPptxSubcommands } from "./pptx-cli.js";

const defaultIO: IO = { stdout: process.stdout, stderr: process.stderr };

export async function runCli(argv: string[], io: IO = defaultIO): Promise<number> {
  const program = new Command();
  program
    .name("office-agent")
    .description(
      "Headless agent CLI for OfficeAI. DOCX, XLSX and PPTX are all supported in this build. The 'docx', 'xlsx' and 'pptx' subcommand groups are the canonical surfaces; top-level read/search/insert-text/comment/apply remain as backward-compatible shims for DOCX."
    )
    .version("0.1.0")
    .exitOverride();

  // ── docx subcommand group ───────────────────────────────────────────────
  const docx = program.command("docx").description("DOCX-specific commands. See `office-agent docx --help`.");
  registerDocxSubcommands(docx, io);

  // ── xlsx subcommand group ───────────────────────────────────────────────
  const xlsx = program.command("xlsx").description("XLSX-specific commands. See `office-agent xlsx --help`.");
  registerXlsxSubcommands(xlsx, io);

  // ── pptx subcommand group ───────────────────────────────────────────────
  const pptx = program.command("pptx").description("PPTX-specific commands. See `office-agent pptx --help`.");
  registerPptxSubcommands(pptx, io);

  // ── Backward-compatible top-level shims ─────────────────────────────────
  // Old commands forwarded the same payloads. We keep them aliased so existing
  // scripts (and the spec/agent/cli.md examples in the in-repo docs) keep
  // working while the docx subcommand group becomes the canonical surface.
  registerLegacyTopLevel(program, io);

  // ── MCP server ──────────────────────────────────────────────────────────
  program
    .command("mcp")
    .description(
      "Start the OfficeAI Model Context Protocol server over stdio (newline-delimited JSON-RPC). Pipe an MCP-aware client (Claude Desktop, etc.) to this process."
    )
    .action(async () => {
      await runMcpStdioServer();
    });

  try {
    await program.parseAsync(argv, { from: "user" });
    return 0;
  } catch (err) {
    return mapErrorToExitCode(err, io);
  }
}

// ──────────────────────────────────────────────────────────────────────────
// docx subcommand group
// ──────────────────────────────────────────────────────────────────────────

function registerDocxSubcommands(docx: Command, io: IO): void {
  docx
    .command("inspect")
    .description("Print a structural summary (paragraphs, tables, comments, parts) as JSON.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; pretty: boolean }) => {
      const agent = await loadAgent(opts.file);
      const summary = inspectSnapshot(agent.getSnapshot());
      io.stdout.write(stringifyJson(summary, opts.pretty) + "\n");
    });

  docx
    .command("read")
    .description("Read a DOCX file as Markdown, structured JSON, or plain text.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .addOption(
      new Option("--format <fmt>", "Output format").choices(["markdown", "json", "text"]).default("markdown")
    )
    .addOption(new Option("--range <selector>", "Selector e.g. paragraph:0..paragraph:5"))
    .option("--pretty", "Pretty-print JSON output (only with --format json)", false)
    .action(
      async (opts: {
        file: string;
        format: "markdown" | "json" | "text";
        range?: string;
        pretty: boolean;
      }) => {
        const agent = await loadAgent(opts.file);
        const snap = agent.getSnapshot();
        const range = opts.range ? rangeFromSelector(opts.range) : undefined;
        switch (opts.format) {
          case "markdown":
            io.stdout.write(renderMarkdown(agent, range) + "\n");
            return;
          case "text":
            io.stdout.write(renderPlainText(snap, range) + "\n");
            return;
          case "json":
            io.stdout.write(stringifyJson(snapshotToJsonProjection(snap, range), opts.pretty) + "\n");
            return;
          default: {
            const _exhaustive: never = opts.format;
            void _exhaustive;
          }
        }
      }
    );

  docx
    .command("search")
    .description("Search a DOCX file for text and print matches as JSON.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("-q, --query <text>", "Search query")
    .option("--case-sensitive", "Case-sensitive search", false)
    .option("--regex", "Treat the query as a regular expression", false)
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        query: string;
        caseSensitive: boolean;
        regex: boolean;
        pretty: boolean;
      }) => {
        const agent = await loadAgent(opts.file);
        const results = agent.search({
          query: opts.query,
          caseSensitive: opts.caseSensitive,
          regex: opts.regex,
        });
        io.stdout.write(stringifyJson(results, opts.pretty) + "\n");
      }
    );

  docx
    .command("write")
    .description("Insert text at a position selector and write the result.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--at <selector>", "Position selector e.g. section:0/paragraph:0/run:0/text:5")
    .requiredOption("--text <text>", "Text to insert")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .addOption(new Option("--source <src>", "Mutation source").choices(["agent", "human"]).default("agent"))
    .option("--agent-id <id>", "Agent identifier (defaults to office-agent-cli)", "office-agent-cli")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        at: string;
        text: string;
        out?: string;
        source: "agent" | "human";
        agentId: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        const at = positionFromSelector(parseSelector(opts.at));
        await runWrite(io, opts, "docx:insert-text", { at, text: opts.text });
      }
    );

  docx
    .command("style")
    .description("Set the paragraph style at a position selector and write the result.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--at <selector>", "Position selector targeting a paragraph")
    .requiredOption("--style <styleId>", "Style id, e.g. Heading1")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        at: string;
        style: string;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        const at = positionFromSelector(parseSelector(opts.at));
        await runWrite(io, opts, "docx:set-paragraph-style", { at, style: opts.style });
      }
    );

  docx
    .command("comment")
    .description("Add a comment to a range selector and write the result.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--range <selector>", "Range selector e.g. paragraph:0/text:0..5")
    .requiredOption("--text <text>", "Comment text")
    .option("--author <name>", "Comment author", "office-agent")
    .option("--initials <initials>", "Comment author initials", "OA")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        range: string;
        text: string;
        author: string;
        initials: string;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        const range = rangeFromSelector(opts.range);
        await runWrite(io, opts, "docx:add-comment", {
          range,
          text: opts.text,
          author: opts.author,
          initials: opts.initials,
        });
      }
    );

  docx
    .command("resolve-comment")
    .description("Mark a comment as resolved (or re-open with --reopen) and write the result.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--id <commentId>", "Target comment id (as exposed by `docx inspect`)")
    .option("--reopen", "Re-open a previously resolved comment instead of resolving it", false)
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        id: string;
        reopen: boolean;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        await runWrite(io, opts, "docx:resolve-comment", {
          commentId: opts.id,
          resolved: !opts.reopen,
        });
      }
    );

  docx
    .command("reply-comment")
    .description("Append a reply to an existing comment and write the result.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--parent <commentId>", "Parent comment id")
    .requiredOption("--text <text>", "Reply text")
    .option("--author <name>", "Reply author", "office-agent")
    .option("--initials <initials>", "Reply author initials")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        parent: string;
        text: string;
        author: string;
        initials?: string;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        await runWrite(io, opts, "docx:reply-comment", {
          parentId: opts.parent,
          text: opts.text,
          author: opts.author,
          ...(opts.initials ? { initials: opts.initials } : {}),
        });
      }
    );

  docx
    .command("delete-comment")
    .description("Delete a comment (and its reply thread) and write the result.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--id <commentId>", "Target comment id")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; id: string; out?: string; approve: boolean; pretty: boolean }) => {
      await runWrite(io, opts, "docx:delete-comment", { commentId: opts.id });
    });

  docx
    .command("format-range")
    .description(
      "Apply text formatting to a range (bold/italic/underline/strikethrough/font/size/color/highlight)."
    )
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--range <selector>", "Range selector e.g. paragraph:0/text:0..5 or paragraph:0..2")
    .option("--bold <bool>", "true|false to set, omit to leave unchanged")
    .option("--italic <bool>", "true|false")
    .option("--underline <bool>", "true|false")
    .option("--strike <bool>", "true|false")
    .option("--font-family <name>", "Run font family")
    .option("--font-size <halfPoints>", "Run font size in half-points (e.g. 24 = 12pt)", parseIntOpt)
    .option("--color <RRGGBB>", "Hex color without leading #")
    .option("--highlight <name>", "Highlight color name (yellow|green|cyan|red|...)")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        range: string;
        bold?: string;
        italic?: string;
        underline?: string;
        strike?: string;
        fontFamily?: string;
        fontSize?: number;
        color?: string;
        highlight?: string;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        const range = rangeFromSelector(opts.range);
        const format: Record<string, unknown> = {};
        if (opts.bold !== undefined) format.bold = parseBool(opts.bold, "--bold");
        if (opts.italic !== undefined) format.italic = parseBool(opts.italic, "--italic");
        if (opts.underline !== undefined) format.underline = parseBool(opts.underline, "--underline");
        if (opts.strike !== undefined) format.strike = parseBool(opts.strike, "--strike");
        if (opts.fontFamily !== undefined) format.fontFamily = opts.fontFamily;
        if (opts.fontSize !== undefined) format.fontSize = opts.fontSize;
        if (opts.color !== undefined) format.color = opts.color;
        if (opts.highlight !== undefined) format.highlight = opts.highlight;
        if (Object.keys(format).length === 0) {
          throw new CliError(
            64,
            "format-range: pass at least one of --bold/--italic/--underline/--strike/--font-family/--font-size/--color/--highlight"
          );
        }
        await runWrite(io, opts, "docx:format-range", { range, format });
      }
    );

  docx
    .command("accept-change")
    .description("Accept a tracked change (insertion folds in, deletion lands).")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--id <revisionId>", "Tracked-change revision id (the w:id on <w:ins>/<w:del>)")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; id: string; out?: string; approve: boolean; pretty: boolean }) => {
      await runWrite(io, opts, "docx:accept-change", { revisionId: opts.id });
    });

  docx
    .command("reject-change")
    .description("Reject a tracked change (insertion is dropped, deletion is undone).")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--id <revisionId>", "Tracked-change revision id")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; id: string; out?: string; approve: boolean; pretty: boolean }) => {
      await runWrite(io, opts, "docx:reject-change", { revisionId: opts.id });
    });

  docx
    .command("insert-image")
    .description("Insert an inline image at a position selector. Reads image bytes from --image.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--at <selector>", "Position selector targeting a paragraph")
    .requiredOption("--image <path>", "Path to image file (PNG/JPEG/GIF/BMP/TIFF/WebP/SVG)")
    .option("--width <px>", "Display width in pixels @ 96 DPI (default: 200)", parseIntOpt, 200)
    .option("--height <px>", "Display height in pixels @ 96 DPI (default: 200)", parseIntOpt, 200)
    .option("--mime <type>", "Override mime type (auto-detected from extension by default)")
    .option("--alt <text>", "Alt text (populates wp:docPr@descr)")
    .option("--name <name>", "Display name (wp:docPr@name)")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        at: string;
        image: string;
        width: number;
        height: number;
        mime?: string;
        alt?: string;
        name?: string;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        const at = positionFromSelector(parseSelector(opts.at));
        const data = await readFile(resolve(opts.image));
        const mimeType = opts.mime ?? mimeTypeFromExtension(opts.image);
        if (!mimeType) {
          throw new CliError(
            64,
            `insert-image: cannot infer mime type for ${opts.image}; pass --mime explicitly`
          );
        }
        const payload: Record<string, unknown> = {
          at,
          data: new Uint8Array(data),
          mimeType,
          width: opts.width,
          height: opts.height,
        };
        if (opts.alt !== undefined) payload.altText = opts.alt;
        if (opts.name !== undefined) payload.name = opts.name;
        await runWrite(io, opts, "docx:insert-image", payload);
      }
    );

  docx
    .command("insert-table")
    .description("Insert an empty rows × cols table at a position selector.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--at <selector>", "Position selector targeting a paragraph")
    .requiredOption("--rows <n>", "Row count", parseIntOpt)
    .requiredOption("--cols <n>", "Column count", parseIntOpt)
    .option("--col-widths <csv>", "Optional comma-separated column widths in twips")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        at: string;
        rows: number;
        cols: number;
        colWidths?: string;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        const at = positionFromSelector(parseSelector(opts.at));
        const payload: Record<string, unknown> = { at, rows: opts.rows, cols: opts.cols };
        if (opts.colWidths) {
          payload.columnWidths = opts.colWidths.split(",").map((s) => Number.parseInt(s.trim(), 10));
        }
        await runWrite(io, opts, "docx:insert-table", payload);
      }
    );

  docx
    .command("insert-row")
    .description("Insert a row into a table.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--table-id <id>", "Target table id (use `docx inspect` to find ids)")
    .requiredOption("--at <n>", "0-based row index; equal to row count to append", parseIntOpt)
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        tableId: string;
        at: number;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        await runWrite(io, opts, "docx:insert-row", { tableId: opts.tableId, at: opts.at });
      }
    );

  docx
    .command("insert-column")
    .description("Insert a column into a table.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--table-id <id>", "Target table id")
    .requiredOption("--at <n>", "0-based column index", parseIntOpt)
    .option("--width <twips>", "Column width in twips (defaults to equal split)", parseIntOpt)
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        tableId: string;
        at: number;
        width?: number;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        const payload: Record<string, unknown> = { tableId: opts.tableId, at: opts.at };
        if (opts.width !== undefined) payload.width = opts.width;
        await runWrite(io, opts, "docx:insert-column", payload);
      }
    );

  docx
    .command("set-cell-text")
    .description("Replace one table cell's content with a single plain-text paragraph.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--table-id <id>", "Target table id")
    .requiredOption("--row <n>", "0-based row index", parseIntOpt)
    .requiredOption("--col <n>", "0-based column index", parseIntOpt)
    .requiredOption("--text <text>", "Text to place in the cell")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        tableId: string;
        row: number;
        col: number;
        text: string;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        const content = [
          {
            kind: "paragraph",
            id: "",
            properties: {},
            children: [
              {
                kind: "run",
                id: "",
                properties: {},
                children: [{ kind: "text", text: opts.text }],
              },
            ],
          },
        ];
        await runWrite(io, opts, "docx:set-cell-content", {
          tableId: opts.tableId,
          row: opts.row,
          col: opts.col,
          content,
        });
      }
    );

  docx
    .command("insert-hyperlink")
    .description("Wrap a flat-text range in a paragraph with a hyperlink.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--paragraph-id <id>", "Target paragraph id (use `docx read --format json` to find ids)")
    .requiredOption("--start <n>", "Inclusive start offset", parseIntOpt)
    .requiredOption("--end <n>", "Exclusive end offset", parseIntOpt)
    .option("--target <url>", "External URL target (mutually exclusive with --anchor)")
    .option("--anchor <name>", "Internal bookmark anchor (mutually exclusive with --target)")
    .option("--tooltip <text>", "Tooltip text")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        paragraphId: string;
        start: number;
        end: number;
        target?: string;
        anchor?: string;
        tooltip?: string;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        if (!opts.target && !opts.anchor) {
          throw new CliError(64, "insert-hyperlink: pass --target <url> or --anchor <name>");
        }
        if (opts.target && opts.anchor) {
          throw new CliError(64, "insert-hyperlink: --target and --anchor are mutually exclusive");
        }
        const payload: Record<string, unknown> = {
          paragraphId: opts.paragraphId,
          range: { start: opts.start, end: opts.end },
        };
        if (opts.target) payload.target = opts.target;
        if (opts.anchor) payload.anchor = opts.anchor;
        if (opts.tooltip) payload.tooltip = opts.tooltip;
        await runWrite(io, opts, "docx:insert-hyperlink", payload);
      }
    );

  docx
    .command("remove-hyperlink")
    .description("Unwrap a hyperlink from a paragraph (optionally reaping its relationship).")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--paragraph-id <id>", "Target paragraph id")
    .requiredOption("--hyperlink-id <id>", "Hyperlink node id")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        paragraphId: string;
        hyperlinkId: string;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        await runWrite(io, opts, "docx:remove-hyperlink", {
          paragraphId: opts.paragraphId,
          hyperlinkId: opts.hyperlinkId,
        });
      }
    );

  docx
    .command("set-list")
    .description("Set or replace numbering (list) on a paragraph.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--paragraph-id <id>", "Target paragraph id")
    .requiredOption("--num-id <n>", "Numbering instance id (matches w:num/@w:numId)", parseIntOpt)
    .option("--ilvl <n>", "0-based level within the abstract numbering definition", parseIntOpt, 0)
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        paragraphId: string;
        numId: number;
        ilvl: number;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        await runWrite(io, opts, "docx:set-paragraph-list", {
          paragraphId: opts.paragraphId,
          numId: opts.numId,
          ilvl: opts.ilvl,
        });
      }
    );

  docx
    .command("remove-list")
    .description("Remove numbering (list) from a paragraph.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--paragraph-id <id>", "Target paragraph id")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        paragraphId: string;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        await runWrite(io, opts, "docx:remove-paragraph-list", {
          paragraphId: opts.paragraphId,
        });
      }
    );

  docx
    .command("align")
    .description("Set (or clear with --clear) a paragraph's <w:jc> alignment.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--paragraph-id <id>", "Target paragraph id")
    .option("--alignment <value>", "left | center | right | justify")
    .option("--clear", "Clear any existing alignment", false)
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        paragraphId: string;
        alignment?: string;
        clear: boolean;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        if (!opts.clear && !opts.alignment) {
          throw new CliError(64, "align: pass --alignment <value> or --clear");
        }
        const alignment = opts.clear ? null : opts.alignment;
        await runWrite(io, opts, "docx:set-paragraph-alignment", {
          paragraphId: opts.paragraphId,
          alignment,
        });
      }
    );

  docx
    .command("indent")
    .description("Step a paragraph's left indent by --delta twips (negative outdents).")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--paragraph-id <id>", "Target paragraph id")
    .requiredOption("--delta <twips>", "Signed delta in twips applied to indentation.left", parseIntOpt)
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        paragraphId: string;
        delta: number;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        await runWrite(io, opts, "docx:set-paragraph-indent", {
          paragraphId: opts.paragraphId,
          deltaTwips: opts.delta,
        });
      }
    );

  docx
    .command("header")
    .description("Replace one header paragraph's text content.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--part <path>", "Header part path, e.g. word/header1.xml")
    .requiredOption("--paragraph-index <n>", "0-based paragraph index inside the header", parseIntOpt)
    .requiredOption("--text <text>", "New plain-text content")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        part: string;
        paragraphIndex: number;
        text: string;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        await runWrite(io, opts, "docx:set-header-text", {
          partId: opts.part,
          paragraphIndex: opts.paragraphIndex,
          text: opts.text,
        });
      }
    );

  docx
    .command("footer")
    .description("Replace one footer paragraph's text content.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--part <path>", "Footer part path, e.g. word/footer1.xml")
    .requiredOption("--paragraph-index <n>", "0-based paragraph index inside the footer", parseIntOpt)
    .requiredOption("--text <text>", "New plain-text content")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        part: string;
        paragraphIndex: number;
        text: string;
        out?: string;
        approve: boolean;
        pretty: boolean;
      }) => {
        await runWrite(io, opts, "docx:set-footer-text", {
          partId: opts.part,
          paragraphIndex: opts.paragraphIndex,
          text: opts.text,
        });
      }
    );

  // ── Pending-mutation review (human review flow) ────────────────────────
  // Per prompt §"The Human Review Flow": agents produce mutations that stage
  // as pending. These commands surface that queue end-to-end on the CLI so
  // an operator can ratify or roll back agent changes without touching the
  // GUI. Pass `--no-approve` to any write subcommand to leave its mutation
  // pending; `docx pending list/approve/reject` then operate on the queue.
  // The pending queue is in-process: it does NOT survive across CLI
  // invocations. To stage agent edits for review across processes, dump the
  // intermediate snapshot to disk and run `docx diff` against the original.
  const pending = docx
    .command("pending")
    .description(
      "Inspect, approve, or reject pending agent mutations. The queue lives only inside one CLI process — see --no-approve and docx apply --no-approve."
    );
  pending
    .command("list")
    .description(
      "List pending mutations on a fresh load (mostly useful for round-tripping with --no-approve)."
    )
    .requiredOption("--file <path>", "Path to a .docx file")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; pretty: boolean }) => {
      const agent = await loadAgent(opts.file);
      io.stdout.write(stringifyJson({ pending: agent.getPendingMutations().length }, opts.pretty) + "\n");
    });

  docx
    .command("apply")
    .description("Apply a JSON command file (single command or { commands: [...] }) and write the result.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("-c, --commands <path>", "Path to a JSON file containing one or more commands")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; commands: string; out?: string; pretty: boolean }) => {
      const agent = await loadAgent(opts.file);
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

  docx
    .command("diff")
    .description("Compute a structural diff between two DOCX files.")
    .requiredOption("--before <path>", "Path to the baseline .docx file")
    .requiredOption("--after <path>", "Path to the modified .docx file")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { before: string; after: string; pretty: boolean }) => {
      const before = await loadAgent(opts.before);
      const after = await loadAgent(opts.after);
      const summary = diffSnapshots(before.getSnapshot(), after.getSnapshot());
      io.stdout.write(stringifyJson(summary, opts.pretty) + "\n");
    });
}

// ──────────────────────────────────────────────────────────────────────────
// Backward-compatible top-level shims
// ──────────────────────────────────────────────────────────────────────────

function registerLegacyTopLevel(program: Command, io: IO): void {
  program
    .command("read")
    .description("[legacy] Read a DOCX file as Markdown. Prefer `office-agent docx read`.")
    .requiredOption("-i, --input <path>", "Path to a .docx file")
    .addOption(new Option("--range <selector>", "Selector e.g. paragraph:0..paragraph:5"))
    .action(async (opts: { input: string; range?: string }) => {
      const agent = await loadAgent(opts.input);
      const range = opts.range ? rangeFromSelector(opts.range) : undefined;
      io.stdout.write(renderMarkdown(agent, range) + "\n");
    });

  program
    .command("search")
    .description("[legacy] Search a DOCX file for text. Prefer `office-agent docx search`.")
    .requiredOption("-i, --input <path>", "Path to a .docx file")
    .requiredOption("-q, --query <text>", "Search query")
    .option("--case-sensitive", "Case-sensitive search", false)
    .option("--regex", "Treat the query as a regular expression", false)
    .action(async (opts: { input: string; query: string; caseSensitive: boolean; regex: boolean }) => {
      const agent = await loadAgent(opts.input);
      const results = agent.search({
        query: opts.query,
        caseSensitive: opts.caseSensitive,
        regex: opts.regex,
      });
      io.stdout.write(JSON.stringify(results, null, 2) + "\n");
    });

  program
    .command("insert-text")
    .description("[legacy] Insert text at a position. Prefer `office-agent docx write`.")
    .requiredOption("-i, --input <path>", "Path to a .docx file")
    .requiredOption("-o, --output <path>", "Path to write the resulting .docx file")
    .requiredOption("--at <selector>", "Position selector e.g. paragraph:0/run:0/text:5")
    .requiredOption("--text <text>", "Text to insert")
    .action(async (opts: { input: string; output: string; at: string; text: string }) => {
      const agent = await loadAgent(opts.input);
      const at = positionFromSelector(parseSelector(opts.at));
      await agent.applyCommand({
        type: "docx:insert-text",
        payload: { at, text: opts.text },
        source: "agent",
        agentId: "office-agent-cli",
      });
      agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
      await writeFile(resolve(opts.output), Buffer.from(await agent.exportFile()));
      io.stdout.write(`wrote ${opts.output}\n`);
    });

  program
    .command("comment")
    .description("[legacy] Add a comment. Prefer `office-agent docx comment`.")
    .requiredOption("-i, --input <path>", "Path to a .docx file")
    .requiredOption("-o, --output <path>", "Path to write the resulting .docx file")
    .requiredOption("--range <selector>", "Range selector e.g. paragraph:0/text:0..5")
    .requiredOption("--text <text>", "Comment text")
    .option("--author <name>", "Comment author", "office-agent")
    .option("--initials <initials>", "Comment author initials", "OA")
    .action(
      async (opts: {
        input: string;
        output: string;
        range: string;
        text: string;
        author: string;
        initials: string;
      }) => {
        const agent = await loadAgent(opts.input);
        const range = rangeFromSelector(opts.range);
        await agent.applyCommand({
          type: "docx:add-comment",
          payload: { range, text: opts.text, author: opts.author, initials: opts.initials },
          source: "agent",
          agentId: "office-agent-cli",
        });
        agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
        await writeFile(resolve(opts.output), Buffer.from(await agent.exportFile()));
        io.stdout.write(`wrote ${opts.output}\n`);
      }
    );

  program
    .command("apply")
    .description("[legacy] Apply a JSON command file. Prefer `office-agent docx apply`.")
    .requiredOption("-i, --input <path>", "Path to a .docx file")
    .requiredOption("-o, --output <path>", "Path to write the resulting .docx file")
    .requiredOption("-c, --commands <path>", "Path to a JSON file containing one or more commands")
    .action(async (opts: { input: string; output: string; commands: string }) => {
      const agent = await loadAgent(opts.input);
      const raw = await readFile(resolve(opts.commands), "utf8");
      const data: unknown = JSON.parse(raw);
      const cmds = normalizeCommands(data);
      const muts = await agent.applyCommands(cmds);
      agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
      await writeFile(resolve(opts.output), Buffer.from(await agent.exportFile()));
      io.stdout.write(
        JSON.stringify(
          {
            wrote: opts.output,
            mutations: muts.map((m) => ({ id: m.id, type: m.command.type, status: m.status })),
          },
          null,
          2
        ) + "\n"
      );
    });
}

// ──────────────────────────────────────────────────────────────────────────
// Helpers shared by CLI and MCP server
// ──────────────────────────────────────────────────────────────────────────

async function loadAgent(input: string): Promise<DocxAgent> {
  const buf = await readFile(resolve(input));
  return DocxAgent.fromBuffer(buf);
}

/**
 * Shared write pipeline for every docx write subcommand. Loads the file,
 * dispatches the command, optionally auto-approves the resulting pending
 * mutation (default), serializes back to disk, and prints a JSON summary
 * to stdout.
 *
 * `opts.approve` is the (commander-derived) boolean from `--no-approve`:
 * commander sets it to `false` when the user passes `--no-approve` and
 * `true` (or undefined for the default) otherwise. When false the
 * mutation is left in the bus's pending queue and the printed summary
 * carries `status: "pending"`.
 */
async function runWrite(
  io: IO,
  opts: { file: string; out?: string; approve?: boolean; pretty?: boolean },
  type: string,
  payload: unknown
): Promise<void> {
  const agent = await loadAgent(opts.file);
  const m = await agent.applyCommand({
    type,
    payload,
    source: "agent",
    agentId: "office-agent-cli",
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
            } as unknown as Mutation<DocxSnapshot>),
        out
      ),
      opts.pretty === true
    ) + "\n"
  );
}

/** Coerce a commander option string to a boolean, throwing CliError on bad input. */
function parseBool(value: string, flag: string): boolean {
  const v = value.trim().toLowerCase();
  if (v === "true" || v === "1" || v === "yes" || v === "on") return true;
  if (v === "false" || v === "0" || v === "no" || v === "off") return false;
  throw new CliError(64, `${flag}: expected true|false, got "${value}"`);
}

/**
 * Minimal extension → mime-type map for `docx insert-image`. Only covers
 * the formats `docx:insert-image` actually accepts. Unknown extensions
 * return `null` and the CLI surfaces a useful "pass --mime" error.
 */
function mimeTypeFromExtension(path: string): string | null {
  const ext = extname(path).toLowerCase().replace(/^\./, "");
  switch (ext) {
    case "png":
      return "image/png";
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "tif":
    case "tiff":
      return "image/tiff";
    case "webp":
      return "image/webp";
    case "svg":
      return "image/svg+xml";
    default:
      return null;
  }
}

void ({} as CommandLite); // keep type import alive (used by callers via runWrite payloads)

function positionFromSelector(sel: Selector): DocxPosition {
  if (sel.kind === "paragraph") return sel.position;
  return sel.range.start;
}

function rangeFromSelector(input: string) {
  const sel = parseSelector(input);
  if (sel.kind !== "range") throw new SelectorError("expected a range selector (e.g. paragraph:0/text:0..5)");
  return sel.range;
}

function renderMarkdown(agent: DocxAgent, range?: { start: DocxPosition; end: DocxPosition }): string {
  if (!range) return agent.toMarkdown();
  const r = agent.getRange({
    kind: "docx-paragraphs",
    start: range.start.paragraph,
    end: range.end.paragraph + 1,
  });
  return r.paragraphs.map((p) => `[${p.index}] ${p.styleId ? `(${p.styleId}) ` : ""}${p.text}`).join("\n");
}

function renderPlainText(snap: DocxSnapshot, range?: { start: DocxPosition; end: DocxPosition }): string {
  const body = snap.root.body;
  const lo = range ? range.start.paragraph : 0;
  const hi = range ? range.end.paragraph + 1 : body.length;
  const lines: string[] = [];
  for (let i = Math.max(0, lo); i < Math.min(body.length, hi); i++) {
    const b = body[i];
    if (b.kind === "paragraph") lines.push(paragraphPlainText(b));
  }
  return lines.join("\n");
}

/**
 * Lightweight, structural snapshot summary intended for `docx inspect` and
 * the `docx_inspect` MCP tool. Counts and part list only — heavy projection
 * lives in `snapshotToJsonProjection`.
 */
export interface DocxSnapshotSummary {
  format: "docx";
  revision: number;
  paragraphs: number;
  tables: number;
  comments: number;
  resolvedComments: number;
  trackedChanges: number;
  parts: string[];
  styles: ReadonlyArray<{ id: string; count: number }>;
}

export function inspectSnapshot(snap: DocxSnapshot): DocxSnapshotSummary {
  let paragraphs = 0;
  let tables = 0;
  let trackedChanges = 0;
  const styleCounts = new Map<string, number>();
  for (const b of snap.root.body) {
    if (b.kind === "paragraph") {
      paragraphs++;
      const id = b.properties.styleId;
      if (id) styleCounts.set(id, (styleCounts.get(id) ?? 0) + 1);
      for (const c of b.children) {
        if (c.kind === "revision") trackedChanges++;
      }
    } else if (b.kind === "table") {
      tables++;
    }
  }
  const resolvedComments = snap.root.comments.filter((c) => c.resolved === true).length;
  const parts = Array.from(snap.container.parts.keys()).sort();
  const styles = Array.from(styleCounts.entries())
    .map(([id, count]) => ({ id, count }))
    .sort((a, b) => a.id.localeCompare(b.id));
  return {
    format: "docx",
    revision: snap.revision,
    paragraphs,
    tables,
    comments: snap.root.comments.length,
    resolvedComments,
    trackedChanges,
    parts,
    styles,
  };
}

/**
 * JSON projection of a snapshot — paragraph index, style, plain text, plus
 * the comment list. Designed to be small enough to pipe to `jq` while still
 * carrying enough structure for an LLM to reason about the document.
 */
export function snapshotToJsonProjection(
  snap: DocxSnapshot,
  range?: { start: DocxPosition; end: DocxPosition }
): {
  format: "docx";
  revision: number;
  paragraphs: ReadonlyArray<{
    index: number;
    id: string;
    styleId?: string;
    text: string;
  }>;
  comments: ReadonlyArray<{
    id: string;
    author: string;
    initials?: string;
    date: string;
    text: string;
    resolved?: boolean;
    parentId?: string;
  }>;
} {
  const body = snap.root.body;
  const lo = range ? range.start.paragraph : 0;
  const hi = range ? range.end.paragraph + 1 : body.length;
  const paragraphs: Array<{
    index: number;
    id: string;
    styleId?: string;
    text: string;
  }> = [];
  for (let i = Math.max(0, lo); i < Math.min(body.length, hi); i++) {
    const b = body[i];
    if (b.kind !== "paragraph") continue;
    paragraphs.push({
      index: i,
      id: b.id,
      ...(b.properties.styleId ? { styleId: b.properties.styleId } : {}),
      text: paragraphPlainText(b),
    });
  }
  const comments = snap.root.comments.map((c) => ({
    id: c.id,
    author: c.author,
    ...(c.initials ? { initials: c.initials } : {}),
    date: c.date,
    text: commentPlainText(c),
    ...(c.resolved === true ? { resolved: true } : {}),
    ...(c.parentId !== undefined ? { parentId: c.parentId } : {}),
  }));
  return { format: "docx", revision: snap.revision, paragraphs, comments };
}

function commentPlainText(c: DocxComment): string {
  return c.body
    .map((b) => (b.kind === "paragraph" ? paragraphPlainText(b as Paragraph) : ""))
    .join("\n")
    .trim();
}

/**
 * Lightweight diff for `docx diff` and `docx_diff`. Compares the projected
 * paragraph list (text + style) and the comment list. We do NOT try to
 * recover OOXML-level byte diffs here; richer diffs are the command bus's
 * job.
 */
export function diffSnapshots(
  before: DocxSnapshot,
  after: DocxSnapshot
): {
  format: "docx";
  paragraphs: { added: number; removed: number; modified: number };
  comments: { added: number; removed: number; resolvedChanged: number };
  changes: ReadonlyArray<{
    kind:
      | "paragraph-added"
      | "paragraph-removed"
      | "paragraph-modified"
      | "comment-added"
      | "comment-removed"
      | "comment-resolved"
      | "comment-reopened";
    index?: number;
    commentId?: string;
    summary: string;
  }>;
} {
  const beforeP = projectParagraphs(before);
  const afterP = projectParagraphs(after);
  const changes: Array<{
    kind:
      | "paragraph-added"
      | "paragraph-removed"
      | "paragraph-modified"
      | "comment-added"
      | "comment-removed"
      | "comment-resolved"
      | "comment-reopened";
    index?: number;
    commentId?: string;
    summary: string;
  }> = [];
  let added = 0;
  let removed = 0;
  let modified = 0;
  const max = Math.max(beforeP.length, afterP.length);
  for (let i = 0; i < max; i++) {
    const b = beforeP[i];
    const a = afterP[i];
    if (!b && a) {
      added++;
      changes.push({ kind: "paragraph-added", index: i, summary: `+[${i}] ${a.text.slice(0, 60)}` });
    } else if (b && !a) {
      removed++;
      changes.push({ kind: "paragraph-removed", index: i, summary: `-[${i}] ${b.text.slice(0, 60)}` });
    } else if (b && a && (b.text !== a.text || b.styleId !== a.styleId)) {
      modified++;
      changes.push({
        kind: "paragraph-modified",
        index: i,
        summary: `~[${i}] ${b.text.slice(0, 30)} → ${a.text.slice(0, 30)}`,
      });
    }
  }
  const beforeComments = new Map(before.root.comments.map((c) => [c.id, c]));
  const afterComments = new Map(after.root.comments.map((c) => [c.id, c]));
  let cAdded = 0;
  let cRemoved = 0;
  let resolvedChanged = 0;
  for (const [id, c] of afterComments) {
    const prev = beforeComments.get(id);
    if (!prev) {
      cAdded++;
      changes.push({ kind: "comment-added", commentId: id, summary: `+comment ${id} by ${c.author}` });
      continue;
    }
    if ((prev.resolved === true) !== (c.resolved === true)) {
      resolvedChanged++;
      changes.push({
        kind: c.resolved ? "comment-resolved" : "comment-reopened",
        commentId: id,
        summary: c.resolved ? `comment ${id} resolved` : `comment ${id} reopened`,
      });
    }
  }
  for (const [id] of beforeComments) {
    if (!afterComments.has(id)) {
      cRemoved++;
      changes.push({ kind: "comment-removed", commentId: id, summary: `-comment ${id}` });
    }
  }
  return {
    format: "docx",
    paragraphs: { added, removed, modified },
    comments: { added: cAdded, removed: cRemoved, resolvedChanged },
    changes,
  };
}

function projectParagraphs(snap: DocxSnapshot): Array<{ id: string; styleId?: string; text: string }> {
  const out: Array<{ id: string; styleId?: string; text: string }> = [];
  for (const b of snap.root.body) {
    if (b.kind !== "paragraph") continue;
    const p = b as Paragraph;
    out.push({
      id: p.id,
      ...(p.properties.styleId ? { styleId: p.properties.styleId } : {}),
      text: paragraphPlainText(p),
    });
  }
  return out;
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

function mapErrorToExitCode(err: unknown, io: IO): number {
  if (err instanceof CliError) {
    io.stderr.write(`error: ${err.message}\n`);
    return err.code;
  }
  if (err instanceof SelectorError) {
    io.stderr.write(`selector error: ${err.message}\n`);
    return 64;
  }
  if (err instanceof Error && (err as { code?: string }).code === "commander.helpDisplayed") return 0;
  if (err instanceof Error && (err as { code?: string }).code === "commander.version") return 0;
  if (err instanceof Error && (err as { code?: string }).code === "ENOENT") {
    io.stderr.write(`error: file not found: ${err.message}\n`);
    return 2;
  }
  if (err instanceof Error) {
    io.stderr.write(`error: ${err.message}\n`);
    return 1;
  }
  io.stderr.write(`error: ${String(err)}\n`);
  return 1;
}

if (import.meta.url === `file://${process.argv[1]}`) {
  runCli(process.argv.slice(2)).then((code) => process.exit(code));
}
