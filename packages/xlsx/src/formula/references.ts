import { ErrorKinds, refWithDeletedTarget, type CellError } from "./errors.js";

/**
 * References — A1 ↔ internal {row, col, abs} translation, sheet-prefix
 * normalisation, and absolute-ref-aware insert/delete adjustment.
 *
 * Internal indices are 0-based; A1 strings are 1-based. The engine
 * never uses R1C1 internally — it's display-only.
 *
 * Spec: `spec/xlsx/formula-engine.md` §8.
 */

/**
 * Bitset describing which dimensions of a ref are absolute (`$`-prefixed).
 *
 * - 0 = neither (relative `A1`)
 * - 1 = row absolute (`A$1`)
 * - 2 = column absolute (`$A1`)
 * - 3 = both (`$A$1`)
 */
export type AbsoluteRef = 0 | 1 | 2 | 3;

export const AbsRef = {
  NONE: 0,
  ROW: 1,
  COLUMN: 2,
  ALL: 3,
} as const satisfies Record<string, AbsoluteRef>;

export interface CellRef {
  readonly sheet: string;
  readonly row: number;
  readonly col: number;
  readonly abs: AbsoluteRef;
}

export interface RangeRef {
  readonly sheet: string;
  readonly r0: number;
  readonly c0: number;
  readonly r1: number;
  readonly c1: number;
  readonly abs0: AbsoluteRef;
  readonly abs1: AbsoluteRef;
}

const MAX_ROW = 1048575; //  Excel limit (rows are 1..1048576 in A1)
const MAX_COL = 16383; //  XFD = 16383 0-based

// ── Column letter conversion ──────────────────────────────────────────────

export function colLetterToIndex(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    const c = letters.charCodeAt(i);
    if (c < 0x41 || c > 0x5a) {
      // accept lowercase too
      if (c >= 0x61 && c <= 0x7a) {
        n = n * 26 + (c - 0x61 + 1);
        continue;
      }
      return -1;
    }
    n = n * 26 + (c - 0x41 + 1);
  }
  return n - 1;
}

export function indexToColLetter(idx: number): string {
  if (idx < 0) throw new Error(`negative col index ${idx}`);
  let n = idx + 1;
  let out = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    out = String.fromCharCode(0x41 + r) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

// ── Sheet-prefix normalisation ────────────────────────────────────────────

const SHEET_QUOTED = /^'((?:[^']|'')*)'!/;
const SHEET_BARE = /^([A-Za-z_][A-Za-z0-9_.]*)!/;

/**
 * Strip a sheet prefix off `text` if present. Returns `{ sheet, rest }`
 * with `sheet` unquoted. Doubled `''` inside quotes collapses to `'`.
 * Returns `undefined` if no prefix found.
 */
export function stripSheetPrefix(text: string): { sheet: string; rest: string } | undefined {
  let m = SHEET_QUOTED.exec(text);
  if (m) {
    return { sheet: m[1].replace(/''/g, "'"), rest: text.slice(m[0].length) };
  }
  m = SHEET_BARE.exec(text);
  if (m) {
    return { sheet: m[1], rest: text.slice(m[0].length) };
  }
  return undefined;
}

const SHEET_NAME_NEEDS_QUOTING = /[^A-Za-z0-9_.]/;
function quoteSheetName(name: string): string {
  if (!SHEET_NAME_NEEDS_QUOTING.test(name) && !/^\d/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

// ── Cell parse / serialise ────────────────────────────────────────────────

const CELL_BODY = /^(\$?)([A-Z]{1,3})(\$?)(\d+)$/i;
const RANGE_BODY = /^(\$?)([A-Z]{1,3})(\$?)(\d+):(\$?)([A-Z]{1,3})(\$?)(\d+)$/i;
const COL_RANGE_BODY = /^(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})$/i;
const ROW_RANGE_BODY = /^(\$?)(\d+):(\$?)(\d+)$/i;

export function parseA1(text: string, defaultSheet: string): CellRef | undefined {
  let body = text;
  let sheet = defaultSheet;
  const sp = stripSheetPrefix(text);
  if (sp) {
    sheet = sp.sheet;
    body = sp.rest;
  }
  const m = CELL_BODY.exec(body);
  if (!m) return undefined;
  const col = colLetterToIndex(m[2]);
  const row = parseInt(m[4], 10) - 1;
  if (col < 0 || col > MAX_COL || row < 0 || row > MAX_ROW) return undefined;
  const abs = (m[1] === "$" ? AbsRef.COLUMN : 0) | (m[3] === "$" ? AbsRef.ROW : 0);
  return { sheet, row, col, abs: abs as AbsoluteRef };
}

export function parseA1Range(text: string, defaultSheet: string): RangeRef | undefined {
  let body = text;
  let sheet = defaultSheet;
  const sp = stripSheetPrefix(text);
  if (sp) {
    sheet = sp.sheet;
    body = sp.rest;
  }

  let m = RANGE_BODY.exec(body);
  if (m) {
    const c0 = colLetterToIndex(m[2]);
    const r0 = parseInt(m[4], 10) - 1;
    const c1 = colLetterToIndex(m[6]);
    const r1 = parseInt(m[8], 10) - 1;
    if (c0 < 0 || r0 < 0 || c1 < 0 || r1 < 0 || c0 > MAX_COL || c1 > MAX_COL || r0 > MAX_ROW || r1 > MAX_ROW)
      return undefined;
    const abs0 = ((m[1] === "$" ? AbsRef.COLUMN : 0) | (m[3] === "$" ? AbsRef.ROW : 0)) as AbsoluteRef;
    const abs1 = ((m[5] === "$" ? AbsRef.COLUMN : 0) | (m[7] === "$" ? AbsRef.ROW : 0)) as AbsoluteRef;
    return normaliseRange({ sheet, r0, c0, r1, c1, abs0, abs1 });
  }

  m = COL_RANGE_BODY.exec(body);
  if (m) {
    const c0 = colLetterToIndex(m[2]);
    const c1 = colLetterToIndex(m[4]);
    if (c0 < 0 || c1 < 0) return undefined;
    const abs0 = (m[1] === "$" ? AbsRef.COLUMN : AbsRef.NONE) as AbsoluteRef;
    const abs1 = (m[3] === "$" ? AbsRef.COLUMN : AbsRef.NONE) as AbsoluteRef;
    return normaliseRange({ sheet, r0: 0, c0, r1: MAX_ROW, c1, abs0, abs1 });
  }

  m = ROW_RANGE_BODY.exec(body);
  if (m) {
    const r0 = parseInt(m[2], 10) - 1;
    const r1 = parseInt(m[4], 10) - 1;
    if (r0 < 0 || r1 < 0) return undefined;
    const abs0 = (m[1] === "$" ? AbsRef.ROW : AbsRef.NONE) as AbsoluteRef;
    const abs1 = (m[3] === "$" ? AbsRef.ROW : AbsRef.NONE) as AbsoluteRef;
    return normaliseRange({ sheet, r0, c0: 0, r1, c1: MAX_COL, abs0, abs1 });
  }

  return undefined;
}

function normaliseRange(r: RangeRef): RangeRef {
  const r0 = Math.min(r.r0, r.r1);
  const r1 = Math.max(r.r0, r.r1);
  const c0 = Math.min(r.c0, r.c1);
  const c1 = Math.max(r.c0, r.c1);
  return { ...r, r0, c0, r1, c1 };
}

export function serializeCellRef(ref: CellRef, anchor?: { sheet: string }): string {
  const colPart = (ref.abs & AbsRef.COLUMN ? "$" : "") + indexToColLetter(ref.col);
  const rowPart = (ref.abs & AbsRef.ROW ? "$" : "") + (ref.row + 1);
  const body = `${colPart}${rowPart}`;
  if (anchor && anchor.sheet === ref.sheet) return body;
  return `${quoteSheetName(ref.sheet)}!${body}`;
}

export function serializeRangeRef(ref: RangeRef, anchor?: { sheet: string }): string {
  const a = serializeCellPart(ref.r0, ref.c0, ref.abs0);
  const b = serializeCellPart(ref.r1, ref.c1, ref.abs1);
  const body = `${a}:${b}`;
  if (anchor && anchor.sheet === ref.sheet) return body;
  return `${quoteSheetName(ref.sheet)}!${body}`;
}

function serializeCellPart(row: number, col: number, abs: AbsoluteRef): string {
  return (abs & AbsRef.COLUMN ? "$" : "") + indexToColLetter(col) + (abs & AbsRef.ROW ? "$" : "") + (row + 1);
}

// ── A1 ↔ R1C1 (display-only) ──────────────────────────────────────────────

export function a1ToR1C1(ref: CellRef, anchor: CellRef): string {
  const r = formatR1C1Component(ref.row, anchor.row, !!(ref.abs & AbsRef.ROW), "R");
  const c = formatR1C1Component(ref.col, anchor.col, !!(ref.abs & AbsRef.COLUMN), "C");
  if (ref.sheet === anchor.sheet) return `${r}${c}`;
  return `${quoteSheetName(ref.sheet)}!${r}${c}`;
}

function formatR1C1Component(idx: number, anchorIdx: number, absolute: boolean, prefix: "R" | "C"): string {
  if (absolute) return `${prefix}${idx + 1}`;
  const delta = idx - anchorIdx;
  if (delta === 0) return prefix;
  return `${prefix}[${delta}]`;
}

const R1C1_PARTS = /^([Rr])(?:(\d+)|\[(-?\d+)\])?([Cc])(?:(\d+)|\[(-?\d+)\])?$/;

export function r1c1ToA1(text: string, anchor: CellRef): CellRef | undefined {
  let body = text;
  let sheet = anchor.sheet;
  const sp = stripSheetPrefix(text);
  if (sp) {
    sheet = sp.sheet;
    body = sp.rest;
  }
  const m = R1C1_PARTS.exec(body);
  if (!m) return undefined;
  const rowAbs = m[2] !== undefined;
  const colAbs = m[5] !== undefined;
  const row = rowAbs
    ? parseInt(m[2], 10) - 1
    : m[3] !== undefined
      ? anchor.row + parseInt(m[3], 10)
      : anchor.row;
  const col = colAbs
    ? parseInt(m[5], 10) - 1
    : m[6] !== undefined
      ? anchor.col + parseInt(m[6], 10)
      : anchor.col;
  if (row < 0 || col < 0) return undefined;
  const abs = ((rowAbs ? AbsRef.ROW : 0) | (colAbs ? AbsRef.COLUMN : 0)) as AbsoluteRef;
  return { sheet, row, col, abs };
}

// ── Insert / delete adjustment ────────────────────────────────────────────

/**
 * Per `EC-R1`: insertion shifts refs at or below the insertion point,
 * regardless of the absolute flag. Per `EC-R2`: a delete that consumes
 * the cell a ref points to returns `#REF!` with `meta.deletedRef`.
 */
export function adjustForInsertRow(
  ref: CellRef | RangeRef,
  sheet: string,
  at: number,
  count: number
): CellRef | RangeRef | CellError {
  if (refSheet(ref) !== sheet || count <= 0) return ref;
  if (isCellRef(ref)) {
    if (ref.row < at) return ref;
    return { ...ref, row: ref.row + count };
  }
  let r0 = ref.r0;
  let r1 = ref.r1;
  if (r0 >= at) r0 += count;
  if (r1 >= at) r1 += count;
  return { ...ref, r0, r1 };
}

export function adjustForDeleteRow(
  ref: CellRef | RangeRef,
  sheet: string,
  at: number,
  count: number
): CellRef | RangeRef | CellError {
  if (refSheet(ref) !== sheet || count <= 0) return ref;
  const end = at + count - 1;
  if (isCellRef(ref)) {
    if (ref.row < at) return ref;
    if (ref.row <= end) return refWithDeletedTarget(serializeCellRef(ref));
    return { ...ref, row: ref.row - count };
  }
  if (ref.r1 < at) return ref;
  if (ref.r0 > end) return { ...ref, r0: ref.r0 - count, r1: ref.r1 - count };
  if (ref.r0 >= at && ref.r1 <= end) {
    // entire range consumed
    return refWithDeletedTarget(serializeRangeRef(ref));
  }
  // partial overlap → shrink
  const r0 = ref.r0 < at ? ref.r0 : at;
  const r1 = ref.r1 - Math.min(count, ref.r1 - at + 1);
  if (r1 < r0) return refWithDeletedTarget(serializeRangeRef(ref));
  return { ...ref, r0, r1 };
}

export function adjustForInsertColumn(
  ref: CellRef | RangeRef,
  sheet: string,
  at: number,
  count: number
): CellRef | RangeRef | CellError {
  if (refSheet(ref) !== sheet || count <= 0) return ref;
  if (isCellRef(ref)) {
    if (ref.col < at) return ref;
    return { ...ref, col: ref.col + count };
  }
  let c0 = ref.c0;
  let c1 = ref.c1;
  if (c0 >= at) c0 += count;
  if (c1 >= at) c1 += count;
  return { ...ref, c0, c1 };
}

export function adjustForDeleteColumn(
  ref: CellRef | RangeRef,
  sheet: string,
  at: number,
  count: number
): CellRef | RangeRef | CellError {
  if (refSheet(ref) !== sheet || count <= 0) return ref;
  const end = at + count - 1;
  if (isCellRef(ref)) {
    if (ref.col < at) return ref;
    if (ref.col <= end) return refWithDeletedTarget(serializeCellRef(ref));
    return { ...ref, col: ref.col - count };
  }
  if (ref.c1 < at) return ref;
  if (ref.c0 > end) return { ...ref, c0: ref.c0 - count, c1: ref.c1 - count };
  if (ref.c0 >= at && ref.c1 <= end) {
    return refWithDeletedTarget(serializeRangeRef(ref));
  }
  const c0 = ref.c0 < at ? ref.c0 : at;
  const c1 = ref.c1 - Math.min(count, ref.c1 - at + 1);
  if (c1 < c0) return refWithDeletedTarget(serializeRangeRef(ref));
  return { ...ref, c0, c1 };
}

function isCellRef(r: CellRef | RangeRef): r is CellRef {
  return "row" in r;
}

function refSheet(r: CellRef | RangeRef): string {
  return r.sheet;
}

export function isRefError(v: CellRef | RangeRef | CellError): v is CellError {
  return "kind" in v && v.kind === ErrorKinds.REF;
}

// ── Cell key helpers (for the dependency graph) ───────────────────────────

export type CellKey = string; // `${sheet}!${row}:${col}`

export function cellRefKey(ref: CellRef): CellKey {
  return `${ref.sheet}!${ref.row}:${ref.col}`;
}

export function makeCellKey(sheet: string, row: number, col: number): CellKey {
  return `${sheet}!${row}:${col}`;
}

const CELL_KEY = /^(.+)!(\d+):(\d+)$/;
export function parseCellKey(key: CellKey): { sheet: string; row: number; col: number } {
  const m = CELL_KEY.exec(key);
  if (!m) throw new Error(`invalid cell key: ${key}`);
  return { sheet: m[1], row: parseInt(m[2], 10), col: parseInt(m[3], 10) };
}
