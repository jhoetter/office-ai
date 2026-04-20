/**
 * Shared utilities used by both the docx and xlsx CLI subcommand groups.
 * Kept in its own module to avoid a circular `cli.ts` ↔ `cli-xlsx.ts`
 * import.
 */

export interface IO {
  stdout: NodeJS.WritableStream;
  stderr: NodeJS.WritableStream;
}

export class CliError extends Error {
  /**
   * When true the top-level error mapper in `cli.ts` skips its
   * `error: <msg>` prefix line. PDF subcommands rely on this so they
   * can emit a structured `{ "error": "<code>", "message": "..." }`
   * envelope to stderr themselves while still returning a non-zero
   * exit code via commander's preAction hook.
   */
  public readonly silent: boolean;

  constructor(code: number, message: string);
  constructor(code: number, message: string, opts: { silent?: boolean });
  constructor(
    public readonly code: number,
    message: string,
    opts: { silent?: boolean } = {}
  ) {
    super(message);
    this.silent = opts.silent === true;
  }
}

export function stringifyJson(value: unknown, pretty: boolean): string {
  return pretty ? JSON.stringify(value, null, 2) : JSON.stringify(value);
}

/** commander option parser that coerces a flag to an integer (or rejects). */
export function parseIntOpt(value: string, _previous?: number): number {
  const n = Number.parseInt(value, 10);
  if (!Number.isFinite(n)) throw new CliError(64, `expected integer, got "${value}"`);
  return n;
}

/**
 * Process-wide flag toggled by `--deterministic-ids` (set on the root
 * `office-agent` Command in `cli.ts`) or the legacy
 * `OFFICEAI_DETERMINISTIC_IDS=1` env var. Loaders in cli.ts /
 * cli-xlsx.ts / pptx-cli.ts read this via `useDeterministicIds()` and
 * pass `deterministicIdMinter()` into the agent factory so script
 * runs against the same fixture mint identical NodeIds across
 * invocations.
 */
let _deterministicIds = false;

export function setDeterministicIds(on: boolean): void {
  _deterministicIds = on;
}

export function useDeterministicIds(): boolean {
  return _deterministicIds || process.env.OFFICEAI_DETERMINISTIC_IDS === "1";
}

/* ── PPTX unit conversion (G6) ──────────────────────────────────────────
 * EMU is the OOXML wire unit but agents reason in pixels/inches/cm/pt.
 * Conversion table per ECMA-376 Part 1, §17.18.93:
 *   1 in   = 914400 EMU
 *   1 pt   =  12700 EMU  (72pt = 1in)
 *   1 cm   = 360000 EMU
 *   1 px   =   9525 EMU  (96 DPI canonical)
 */
const EMU_PER_UNIT: Readonly<Record<EmuUnit, number>> = {
  emu: 1,
  px: 9525,
  in: 914400,
  cm: 360000,
  pt: 12700,
};

export type EmuUnit = "emu" | "px" | "in" | "cm" | "pt";

export const EMU_UNITS: ReadonlyArray<EmuUnit> = ["emu", "px", "in", "cm", "pt"];

export function isEmuUnit(s: string): s is EmuUnit {
  return (EMU_UNITS as ReadonlyArray<string>).includes(s);
}

/**
 * Convert a number in `unit` to EMU (rounded). Validates `unit` and
 * surfaces a typed CliError on bad input so the parent CLI can map it
 * to exit code 64.
 */
export function toEmu(value: number, unit: EmuUnit): number {
  if (!Number.isFinite(value)) throw new CliError(64, `--unit value must be finite, got ${value}`);
  return Math.round(value * EMU_PER_UNIT[unit]);
}

/**
 * Drain `process.stdin` (or any Readable) into a UTF-8 string. Used by
 * the `--from-stdin` and `--payload-stdin` flags on the `apply`
 * subcommands so large/generated command sets can be piped directly
 * into the CLI without touching disk:
 *
 *   echo '{"commands":[…]}' | office-agent docx apply --file in.docx --from-stdin
 *
 * Throws a typed CliError (exit code 64) if stdin is a TTY (interactive
 * shells would block forever waiting for input).
 */
export async function readStdinToString(stdin: NodeJS.ReadableStream = process.stdin): Promise<string> {
  if ((stdin as NodeJS.ReadStream).isTTY) {
    throw new CliError(
      64,
      "expected JSON on stdin but stdin is a TTY; pipe data in (e.g. `cat cmds.json | office-agent …`) or drop the --from-stdin/--payload-stdin flag"
    );
  }
  const chunks: Buffer[] = [];
  for await (const chunk of stdin as AsyncIterable<Buffer | string>) {
    chunks.push(typeof chunk === "string" ? Buffer.from(chunk, "utf8") : chunk);
  }
  return Buffer.concat(chunks).toString("utf8");
}
