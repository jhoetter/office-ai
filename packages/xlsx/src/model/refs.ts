/**
 * A1-notation utilities. All `row` / `col` coordinates in the typed
 * model are 0-based; A1 strings (`"A1"`, `"B2:C7"`) follow Excel's
 * 1-based row + letter-column convention. Convert at the boundary —
 * never inside the model.
 */

const COLUMN_MAX = 16384; // Excel hard limit (XFD)
const ROW_MAX = 1_048_576; // Excel hard limit

export interface CellAddress {
  /** 0-based row. */
  readonly row: number;
  /** 0-based column. */
  readonly col: number;
}

export interface CellRange {
  /** Top-left, 0-based inclusive. */
  readonly start: CellAddress;
  /** Bottom-right, 0-based inclusive. */
  readonly end: CellAddress;
}

/** Stable map key for a typed `Sheet.cells` entry. */
export function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

export function parseCellKey(key: string): CellAddress {
  const idx = key.indexOf(":");
  if (idx === -1) throw new Error(`bad cell key: ${key}`);
  const row = Number(key.slice(0, idx));
  const col = Number(key.slice(idx + 1));
  if (!Number.isInteger(row) || !Number.isInteger(col)) {
    throw new Error(`bad cell key: ${key}`);
  }
  return { row, col };
}

/** Convert a 0-based column index to an Excel letter (`0` → `"A"`, `26` → `"AA"`). */
export function colToLetter(col: number): string {
  if (!Number.isInteger(col) || col < 0 || col >= COLUMN_MAX) {
    throw new Error(`column out of range: ${col}`);
  }
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Convert an Excel letter to a 0-based column index (`"A"` → `0`, `"AA"` → `26`). */
export function letterToCol(letter: string): number {
  if (!letter) throw new Error("empty column letter");
  let n = 0;
  for (const ch of letter) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) throw new Error(`bad column letter: ${letter}`);
    n = n * 26 + (code - 64);
  }
  const col = n - 1;
  if (col < 0 || col >= COLUMN_MAX) throw new Error(`column out of range: ${letter}`);
  return col;
}

const A1_RE = /^\$?([A-Z]+)\$?([1-9][0-9]*)$/;

/** Parse a single-cell A1 string (`"A1"`, `"$B$2"`) to a 0-based address. */
export function parseA1(ref: string): CellAddress {
  const m = A1_RE.exec(ref.trim().toUpperCase());
  if (!m) throw new Error(`invalid A1 reference: ${ref}`);
  const col = letterToCol(m[1]);
  const row = Number(m[2]) - 1;
  if (row < 0 || row >= ROW_MAX) throw new Error(`row out of range: ${ref}`);
  return { row, col };
}

export function formatA1(addr: CellAddress): string {
  return `${colToLetter(addr.col)}${addr.row + 1}`;
}

/** Parse an A1 range (`"A1:B5"`, `"A1"`) to a normalized 0-based rectangle. */
export function parseRange(ref: string): CellRange {
  const trimmed = ref.trim();
  const colon = trimmed.indexOf(":");
  if (colon === -1) {
    const a = parseA1(trimmed);
    return { start: a, end: a };
  }
  const a = parseA1(trimmed.slice(0, colon));
  const b = parseA1(trimmed.slice(colon + 1));
  return {
    start: { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) },
    end: { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) },
  };
}

export function formatRange(range: CellRange): string {
  if (range.start.row === range.end.row && range.start.col === range.end.col) {
    return formatA1(range.start);
  }
  return `${formatA1(range.start)}:${formatA1(range.end)}`;
}

export function rangeArea(range: CellRange): number {
  return (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
}

export function rangesOverlap(a: CellRange, b: CellRange): boolean {
  return !(
    a.end.row < b.start.row ||
    b.end.row < a.start.row ||
    a.end.col < b.start.col ||
    b.end.col < a.start.col
  );
}
