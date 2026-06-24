/**
 * Action-catalogue → commander adapter.
 *
 * Exposes `registerActionsAsSubcommands(parent, actions, ctx)` which
 * walks an `ActionDescriptor[]` and registers a commander subcommand
 * for every entry whose `cliCallable` is true, has a non-null
 * `commandType`, and declares catalogue-owned args/payload metadata.
 *
 * Design intent: the catalogue file in
 * `packages/{format}/src/actions/catalogue.ts` is the single source
 * of truth for CLI flags + label + description. The adapter takes
 * that data and builds:
 *
 *   office-agent <format> <subcommand-name> --file <path> [...args]
 *
 * Each generated subcommand:
 *   1. Loads the file via the supplied `loadAgent` callback.
 *   2. Builds the bus payload via `descriptor.buildPayload(parsedArgs)`.
 *   3. Dispatches the command, optionally auto-approves, and writes
 *      the result back to disk.
 *   4. Prints a JSON summary of the mutation.
 *
 * Subcommands can declare `commandSchema: "custom"` (or omit args /
 * buildPayload) to opt out of auto-generation; the existing
 * hand-rolled commander block in cli.ts / cli-xlsx.ts / pptx-cli.ts /
 * pdf-cli.ts retains ownership for those (typically because they need
 * stdin handling, batch loops, or custom output formats).
 *
 * The adapter keeps zero knowledge of any specific format — it only
 * understands `ActionDescriptor`, the commander API, and a small
 * `AgentDispatchContext` callback bag the host provides for loading
 * / dispatching / writing.
 */

import { Command, Option } from "commander";
import type { ActionArg, ActionDescriptor } from "@officeai/core";
import { CliError, parseIntOpt, stringifyJson, type IO } from "./cli-shared.js";

/**
 * The minimal contract the host (cli.ts / cli-xlsx.ts / …) provides
 * so the generic adapter can load a file, dispatch a command, and
 * write the result without dragging the per-format agent type into
 * this file.
 */
export interface AgentDispatchContext {
  /** Format key, used in error messages. */
  readonly format: string;
  /**
   * Load a document by path and return an opaque agent handle. The
   * handle is forwarded as-is to `dispatchAndWrite`; the adapter
   * never inspects it.
   */
  loadAgent(filePath: string): Promise<unknown>;
  /**
   * Dispatch one command on a previously-loaded agent and write the
   * result. Returns a JSON-serialisable summary that the adapter
   * forwards to stdout.
   */
  dispatchAndWrite(args: {
    readonly agent: unknown;
    readonly filePath: string;
    readonly outPath: string;
    readonly commandType: string;
    readonly payload: unknown;
    readonly source: "agent" | "human";
    readonly agentId: string;
    readonly approve: boolean;
  }): Promise<unknown>;
}

/**
 * Register every catalogue entry that opted into the terminal CLI
 * AND has the metadata needed for auto-generation. Returns the list
 * of action ids that were skipped, so the host can decide whether to
 * fall back to a hand-rolled commander block for them.
 */
export function registerActionsAsSubcommands(
  parent: Command,
  actions: ReadonlyArray<ActionDescriptor>,
  io: IO,
  ctx: AgentDispatchContext
): { registered: ReadonlyArray<string>; skipped: ReadonlyArray<string> } {
  const registered: string[] = [];
  const skipped: string[] = [];

  for (const action of actions) {
    if (!action.cliCallable) continue;
    if (!isCliAutoBindableAction(action)) {
      skipped.push(action.id);
      continue;
    }
    registerOne(parent, action, io, ctx);
    registered.push(action.id);
  }

  return { registered, skipped };
}

export function isCliAutoBindableAction(action: ActionDescriptor): boolean {
  if (!action.cliCallable) return false;
  if (action.hidden) return false;
  if (action.commandType === null || action.commandType === undefined) return false;
  if (action.commandSchema !== "catalogue-args") return false;
  if (!action.args || !action.buildPayload) return false;
  return true;
}

function registerOne(parent: Command, action: ActionDescriptor, io: IO, ctx: AgentDispatchContext): void {
  const subName = subcommandNameFromId(action.id);
  const cmd = parent
    .command(subName)
    .description(action.description)
    .requiredOption("--file <path>", `Path to a .${ctx.format} file`);

  // Add catalogue-declared args.
  const args = action.args ?? [];
  for (const arg of args) {
    addArgToCommand(cmd, arg);
  }

  // Common write flags shared by every mutation subcommand.
  cmd
    .option("--out <path>", `Path to write the resulting .${ctx.format} file (defaults to --file, in place)`)
    .addOption(new Option("--source <src>", "Mutation source").choices(["agent", "human"]).default("agent"))
    .option("--agent-id <id>", "Agent identifier (defaults to office-agent-cli)", "office-agent-cli")
    .option("--no-approve", "Leave the resulting mutation pending instead of auto-approving it")
    .option("--pretty", "Pretty-print JSON output", false);

  cmd.action(async (opts: Record<string, unknown>) => {
    const filePath = opts.file as string;
    const outPath = (opts.out as string | undefined) ?? filePath;
    const source = (opts.source as "agent" | "human" | undefined) ?? "agent";
    const agentId = (opts.agentId as string | undefined) ?? "office-agent-cli";
    const approve = opts.approve !== false;
    const pretty = opts.pretty === true;

    const parsed = collectArgs(opts, args);
    const buildPayload = action.buildPayload;
    if (!buildPayload) {
      throw new CliError(70, `${action.id}: catalogue entry has no buildPayload (internal error)`);
    }
    const payload = buildPayload(parsed);
    const agent = await ctx.loadAgent(filePath);
    const result = await ctx.dispatchAndWrite({
      agent,
      filePath,
      outPath,
      commandType: action.commandType as string,
      payload,
      source,
      agentId,
      approve,
    });
    io.stdout.write(stringifyJson(result, pretty) + "\n");
  });
}

/**
 * Translate a stable action id (`docx.insert-image`) into the
 * commander subcommand name (`insert-image`). The format prefix is
 * dropped because the subcommand is mounted under the format group
 * (`office-agent docx insert-image`).
 */
function subcommandNameFromId(id: string): string {
  const dot = id.indexOf(".");
  return dot === -1 ? id : id.slice(dot + 1);
}

function addArgToCommand(cmd: Command, arg: ActionArg): void {
  switch (arg.kind) {
    case "string":
    case "selector":
    case "filepath":
    case "enum": {
      const opt = new Option(arg.flag, arg.description);
      if (arg.required) opt.makeOptionMandatory(true);
      if (arg.choices) opt.choices(arg.choices as string[]);
      if (arg.default !== undefined) opt.default(arg.default);
      cmd.addOption(opt);
      return;
    }
    case "number": {
      if (arg.required) {
        cmd.requiredOption(arg.flag, arg.description, parseIntOpt);
      } else if (arg.default !== undefined) {
        cmd.option(arg.flag, arg.description, parseIntOpt, arg.default as number);
      } else {
        cmd.option(arg.flag, arg.description, parseIntOpt);
      }
      return;
    }
    case "boolean": {
      cmd.option(arg.flag, arg.description, arg.default === true);
      return;
    }
    case "stringList": {
      cmd.option(arg.flag, arg.description, (val: string, prev?: string[]) => {
        const next = prev ?? [];
        next.push(val);
        return next;
      });
      return;
    }
    default: {
      const _exhaustive: never = arg.kind;
      void _exhaustive;
      throw new Error(`actions-to-cli: unknown arg kind for ${arg.flag}`);
    }
  }
}

/**
 * Pull just the catalogue-declared args (by their `name`) out of the
 * commander-parsed opts bag. Commander stores them under camelCase
 * keys derived from the long-form flag, which we mirror in
 * `ActionArg.name`.
 */
function collectArgs(opts: Record<string, unknown>, args: ReadonlyArray<ActionArg>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const a of args) {
    out[a.name] = opts[a.name];
  }
  return out;
}
