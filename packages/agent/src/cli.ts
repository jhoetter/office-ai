#!/usr/bin/env node
import { readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { Command, Option } from "commander";
import { DocxAgent, paragraphPlainText } from "@officeai/docx";
import type { DocxComment, DocxPosition, DocxSnapshot, Paragraph } from "@officeai/docx";
import { parseSelector, SelectorError, type Selector } from "./selector.js";
import { runMcpStdioServer } from "./mcp.js";

interface IO {
  stdout: NodeJS.WritableStream;
  stderr: NodeJS.WritableStream;
}

const defaultIO: IO = { stdout: process.stdout, stderr: process.stderr };

export async function runCli(argv: string[], io: IO = defaultIO): Promise<number> {
  const program = new Command();
  program
    .name("office-agent")
    .description(
      "Headless agent CLI for OfficeAI. DOCX is supported in this build; XLSX/PPTX commands report 'not yet supported'. The 'docx' subcommand group is the canonical surface; top-level read/search/insert-text/comment/apply remain as backward-compatible shims."
    )
    .version("0.1.0")
    .exitOverride();

  // ── docx subcommand group ───────────────────────────────────────────────
  const docx = program.command("docx").description("DOCX-specific commands. See `office-agent docx --help`.");
  registerDocxSubcommands(docx, io);

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

  // ── XLSX/PPTX deferral stubs ────────────────────────────────────────────
  for (const stub of ["xlsx", "pptx"] as const) {
    const cmd = new Command(stub).description(
      `(stub) ${stub.toUpperCase()} support is deferred to a future session`
    );
    cmd.action(() => {
      io.stderr.write(`${stub.toUpperCase()} support is not yet implemented in office-agent.\n`);
      throw new CliError(2, `${stub} not implemented`);
    });
    program.addCommand(cmd);
  }

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
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        at: string;
        text: string;
        out?: string;
        source: "agent" | "human";
        agentId: string;
        pretty: boolean;
      }) => {
        const agent = await loadAgent(opts.file);
        const at = positionFromSelector(parseSelector(opts.at));
        const m = await agent.applyCommand({
          type: "docx:insert-text",
          payload: { at, text: opts.text },
          source: opts.source,
          ...(opts.source === "agent" ? { agentId: opts.agentId } : {}),
        });
        agent.getPendingMutations().forEach((p) => agent.approveMutation(p.id));
        const out = opts.out ?? opts.file;
        await writeFile(resolve(out), Buffer.from(await agent.exportFile()));
        io.stdout.write(stringifyJson(mutationSummary(m, out), opts.pretty) + "\n");
      }
    );

  docx
    .command("style")
    .description("Set the paragraph style at a position selector and write the result.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--at <selector>", "Position selector targeting a paragraph")
    .requiredOption("--style <styleId>", "Style id, e.g. Heading1")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; at: string; style: string; out?: string; pretty: boolean }) => {
      const agent = await loadAgent(opts.file);
      const at = positionFromSelector(parseSelector(opts.at));
      const m = await agent.applyCommand({
        type: "docx:set-paragraph-style",
        payload: { at, style: opts.style },
        source: "agent",
        agentId: "office-agent-cli",
      });
      agent.getPendingMutations().forEach((p) => agent.approveMutation(p.id));
      const out = opts.out ?? opts.file;
      await writeFile(resolve(out), Buffer.from(await agent.exportFile()));
      io.stdout.write(stringifyJson(mutationSummary(m, out), opts.pretty) + "\n");
    });

  docx
    .command("comment")
    .description("Add a comment to a range selector and write the result.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--range <selector>", "Range selector e.g. paragraph:0/text:0..5")
    .requiredOption("--text <text>", "Comment text")
    .option("--author <name>", "Comment author", "office-agent")
    .option("--initials <initials>", "Comment author initials", "OA")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        range: string;
        text: string;
        author: string;
        initials: string;
        out?: string;
        pretty: boolean;
      }) => {
        const agent = await loadAgent(opts.file);
        const range = rangeFromSelector(opts.range);
        const m = await agent.applyCommand({
          type: "docx:add-comment",
          payload: { range, text: opts.text, author: opts.author, initials: opts.initials },
          source: "agent",
          agentId: "office-agent-cli",
        });
        agent.getPendingMutations().forEach((p) => agent.approveMutation(p.id));
        const out = opts.out ?? opts.file;
        await writeFile(resolve(out), Buffer.from(await agent.exportFile()));
        io.stdout.write(stringifyJson(mutationSummary(m, out), opts.pretty) + "\n");
      }
    );

  docx
    .command("resolve-comment")
    .description("Mark a comment as resolved (or re-open with --reopen) and write the result.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--id <commentId>", "Target comment id (as exposed by `docx inspect`)")
    .option("--reopen", "Re-open a previously resolved comment instead of resolving it", false)
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; id: string; reopen: boolean; out?: string; pretty: boolean }) => {
      const agent = await loadAgent(opts.file);
      const m = await agent.applyCommand({
        type: "docx:resolve-comment",
        payload: { commentId: opts.id, resolved: !opts.reopen },
        source: "agent",
        agentId: "office-agent-cli",
      });
      agent.getPendingMutations().forEach((p) => agent.approveMutation(p.id));
      const out = opts.out ?? opts.file;
      await writeFile(resolve(out), Buffer.from(await agent.exportFile()));
      io.stdout.write(stringifyJson(mutationSummary(m, out), opts.pretty) + "\n");
    });

  docx
    .command("reply-comment")
    .description("Append a reply to an existing comment and write the result.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--parent <commentId>", "Parent comment id")
    .requiredOption("--text <text>", "Reply text")
    .option("--author <name>", "Reply author", "office-agent")
    .option("--initials <initials>", "Reply author initials")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        parent: string;
        text: string;
        author: string;
        initials?: string;
        out?: string;
        pretty: boolean;
      }) => {
        const agent = await loadAgent(opts.file);
        const m = await agent.applyCommand({
          type: "docx:reply-comment",
          payload: {
            parentId: opts.parent,
            text: opts.text,
            author: opts.author,
            ...(opts.initials ? { initials: opts.initials } : {}),
          },
          source: "agent",
          agentId: "office-agent-cli",
        });
        agent.getPendingMutations().forEach((p) => agent.approveMutation(p.id));
        const out = opts.out ?? opts.file;
        await writeFile(resolve(out), Buffer.from(await agent.exportFile()));
        io.stdout.write(stringifyJson(mutationSummary(m, out), opts.pretty) + "\n");
      }
    );

  docx
    .command("delete-comment")
    .description("Delete a comment (and its reply thread) and write the result.")
    .requiredOption("--file <path>", "Path to a .docx file")
    .requiredOption("--id <commentId>", "Target comment id")
    .option("--out <path>", "Path to write the resulting .docx file (defaults to --file, in place)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; id: string; out?: string; pretty: boolean }) => {
      const agent = await loadAgent(opts.file);
      const m = await agent.applyCommand({
        type: "docx:delete-comment",
        payload: { commentId: opts.id },
        source: "agent",
        agentId: "office-agent-cli",
      });
      agent.getPendingMutations().forEach((p) => agent.approveMutation(p.id));
      const out = opts.out ?? opts.file;
      await writeFile(resolve(out), Buffer.from(await agent.exportFile()));
      io.stdout.write(stringifyJson(mutationSummary(m, out), opts.pretty) + "\n");
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

class CliError extends Error {
  constructor(
    public readonly code: number,
    message: string
  ) {
    super(message);
  }
}

async function loadAgent(input: string): Promise<DocxAgent> {
  const buf = await readFile(resolve(input));
  return DocxAgent.fromBuffer(buf);
}

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

function stringifyJson(value: unknown, pretty: boolean): string {
  return pretty ? JSON.stringify(value, null, 2) : JSON.stringify(value);
}

function mapErrorToExitCode(err: unknown, io: IO): number {
  if (err instanceof CliError) return err.code;
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
