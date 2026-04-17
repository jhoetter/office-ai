#!/usr/bin/env node
import { readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { Command, Option } from "commander";
import { DocxAgent } from "@officeai/docx";
import type { DocxPosition } from "@officeai/docx";
import { parseSelector, SelectorError, type Selector } from "./selector.js";

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
      "Headless agent CLI for OfficeAI. DOCX is supported in this build; XLSX/PPTX commands will report 'not yet supported'."
    )
    .version("0.1.0")
    .exitOverride();

  program
    .command("read")
    .description("Read a DOCX file as Markdown (full document or a paragraph range)")
    .requiredOption("-i, --input <path>", "Path to a .docx file")
    .addOption(new Option("--range <selector>", "Selector e.g. paragraph:0..paragraph:5"))
    .action(async (opts: { input: string; range?: string }) => {
      const agent = await loadAgent(opts.input);
      if (opts.range) {
        const sel = parseSelector(opts.range);
        if (sel.kind !== "range") throw new SelectorError("read --range requires a range selector");
        const r = agent.getRange({
          kind: "docx-paragraphs",
          start: sel.range.start.paragraph,
          end: sel.range.end.paragraph + 1,
        });
        for (const p of r.paragraphs) {
          io.stdout.write(`[${p.index}] ${p.styleId ? `(${p.styleId}) ` : ""}${p.text}\n`);
        }
      } else {
        io.stdout.write(agent.toMarkdown() + "\n");
      }
    });

  program
    .command("search")
    .description("Search a DOCX file for text and print matches as JSON")
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
    .description("Insert text at a position selector and write the result")
    .requiredOption("-i, --input <path>", "Path to a .docx file")
    .requiredOption("-o, --output <path>", "Path to write the resulting .docx file")
    .requiredOption("--at <selector>", "Position selector e.g. paragraph:0/run:0/text:5")
    .requiredOption("--text <text>", "Text to insert")
    .action(async (opts: { input: string; output: string; at: string; text: string }) => {
      const agent = await loadAgent(opts.input);
      const sel = parseSelector(opts.at);
      const at = positionFromSelector(sel);
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
    .description("Add a comment to a range selector and write the result")
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
        const sel = parseSelector(opts.range);
        if (sel.kind !== "range") throw new SelectorError("comment --range requires a range selector");
        await agent.applyCommand({
          type: "docx:add-comment",
          payload: {
            range: sel.range,
            text: opts.text,
            author: opts.author,
            initials: opts.initials,
          },
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
    .description(
      "Apply a JSON command file (single command object or { commands: [...] }) and write the result"
    )
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
    if (err instanceof CliError) return err.code;
    if (err instanceof SelectorError) {
      io.stderr.write(`selector error: ${err.message}\n`);
      return 64;
    }
    if (err instanceof Error && (err as { code?: string }).code === "commander.helpDisplayed") return 0;
    if (err instanceof Error && (err as { code?: string }).code === "commander.version") return 0;
    if (err instanceof Error) {
      io.stderr.write(`error: ${err.message}\n`);
      return 1;
    }
    io.stderr.write(`error: ${String(err)}\n`);
    return 1;
  }
}

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

if (import.meta.url === `file://${process.argv[1]}`) {
  runCli(process.argv.slice(2)).then((code) => process.exit(code));
}
