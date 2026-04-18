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
  constructor(
    public readonly code: number,
    message: string
  ) {
    super(message);
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
