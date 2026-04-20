/**
 * Parse an A1-style range string into 0-based row/col bounds.
 *
 * Accepts both single-cell (`B4`) and rectangular (`B4:D7`) refs.
 * Returns `null` for anything else so callers can simply skip
 * malformed presence payloads. Bounds are normalised so `r1 <= r2`
 * and `c1 <= c2` even if the peer published the selection
 * "backwards" (e.g. `D7:B4`).
 *
 * Lives in its own module so the remote-selection overlay
 * (`XlsxRemoteSelectionLayer`) and unit tests can share one
 * implementation without pulling JSX into the test runner.
 */
export interface RangeA1 {
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
}

export function parseRangeA1(range: string): RangeA1 | null {
  const single = /^([A-Z]+)(\d+)$/.exec(range);
  if (single) {
    const c = letterToColIndex(single[1]!);
    const r = Number.parseInt(single[2]!, 10) - 1;
    if (!Number.isFinite(c) || !Number.isFinite(r)) return null;
    return { r1: r, r2: r, c1: c, c2: c };
  }
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(range);
  if (!m) return null;
  const c1 = letterToColIndex(m[1]!);
  const r1 = Number.parseInt(m[2]!, 10) - 1;
  const c2 = letterToColIndex(m[3]!);
  const r2 = Number.parseInt(m[4]!, 10) - 1;
  if (![r1, c1, r2, c2].every(Number.isFinite)) return null;
  return {
    r1: Math.min(r1, r2),
    c1: Math.min(c1, c2),
    r2: Math.max(r1, r2),
    c2: Math.max(c1, c2),
  };
}

export function letterToColIndex(letter: string): number {
  let n = 0;
  for (const ch of letter) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
