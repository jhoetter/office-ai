import { CommandError } from "@officeai/core";
import { parseA1, parseRange, type CellAddress, type CellRange } from "../model/refs.js";
import type { MergedCell, Sheet, XlsxWorkbook } from "../model/types.js";

const SHEET_NAME_FORBIDDEN = /[[\]*?:\\/]/;
const RESERVED_NAMES = new Set(["History".toLowerCase()]);

/** Validate an Excel-rule sheet name. Throws `invalid-name` on failure. */
export function validateSheetName(name: string): void {
  if (typeof name !== "string" || name.length < 1 || name.length > 31) {
    throw new CommandError("invalid-name", `Sheet name must be 1–31 characters; got ${JSON.stringify(name)}`);
  }
  if (SHEET_NAME_FORBIDDEN.test(name)) {
    throw new CommandError(
      "invalid-name",
      `Sheet name "${name}" contains forbidden characters (any of: [ ] * ? : / \\)`
    );
  }
  if (name.startsWith("'") || name.endsWith("'")) {
    throw new CommandError("invalid-name", `Sheet name "${name}" must not start or end with a single quote`);
  }
  if (RESERVED_NAMES.has(name.toLowerCase())) {
    throw new CommandError("invalid-name", `Sheet name "${name}" is reserved by Excel`);
  }
}

/** Throw `duplicate-name` if `name` collides with an existing sheet (case-insensitive). */
export function assertUniqueSheetName(workbook: XlsxWorkbook, name: string, excludeIndex: number = -1): void {
  const lowered = name.toLowerCase();
  for (const sheet of workbook.sheets) {
    if (sheet.index === excludeIndex) continue;
    if (sheet.name.toLowerCase() === lowered) {
      throw new CommandError(
        "duplicate-name",
        `Sheet name "${name}" already exists (case-insensitive); existing: ${workbook.sheets.map((s) => s.name).join(", ")}`
      );
    }
  }
}

/** Resolve a sheet by name; throw `unknown-sheet` if missing. */
export function resolveSheet(workbook: XlsxWorkbook, name: string): Sheet {
  const sheet = workbook.sheets.find((s) => s.name === name);
  if (!sheet) {
    throw new CommandError(
      "unknown-sheet",
      `Sheet "${name}" not found; available: ${workbook.sheets.map((s) => s.name).join(", ")}`
    );
  }
  if (sheet.kind !== "worksheet") {
    throw new CommandError(
      "unknown-sheet",
      `Sheet "${name}" is a ${sheet.kind}; only worksheets accept value/range commands`
    );
  }
  return sheet;
}

/** Parse a single-cell A1 ref; throw `invalid-ref` on failure. */
export function parseCellRef(ref: string): CellAddress {
  try {
    return parseA1(ref);
  } catch (err) {
    throw new CommandError(
      "invalid-ref",
      `Invalid cell ref ${JSON.stringify(ref)}: ${(err as Error).message}`
    );
  }
}

/** Parse an A1 range; throw `invalid-range` on failure. */
export function parseRangeRef(ref: string): CellRange {
  try {
    return parseRange(ref);
  } catch (err) {
    throw new CommandError(
      "invalid-range",
      `Invalid range ${JSON.stringify(ref)}: ${(err as Error).message}`
    );
  }
}

/**
 * Find the merge that contains `addr` (if any). Returns the merge or
 * undefined.
 */
export function findContainingMerge(sheet: Sheet, addr: CellAddress): MergedCell | undefined {
  for (const m of sheet.merges) {
    if (addr.row >= m.r1 && addr.row <= m.r2 && addr.col >= m.c1 && addr.col <= m.c2) return m;
  }
  return undefined;
}

/**
 * Reject when `addr` is a non-anchor cell of a merged region.
 * The anchor is the top-left of the merge.
 */
export function assertNotMergedNonAnchor(sheet: Sheet, addr: CellAddress): void {
  const merge = findContainingMerge(sheet, addr);
  if (!merge) return;
  const isAnchor = merge.r1 === addr.row && merge.c1 === addr.col;
  if (isAnchor) return;
  const anchor = `${columnLetter(merge.c1)}${merge.r1 + 1}`;
  throw new CommandError(
    "merged-non-anchor",
    `Cell is a non-anchor cell of merge ${anchor}:${columnLetter(merge.c2)}${merge.r2 + 1}; target the anchor "${anchor}" or call xlsx:unmerge-cells first`
  );
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

/** Reject if `value` is a string starting with `=`. */
export function assertNotFormulaString(value: unknown): void {
  if (typeof value === "string" && value.startsWith("=")) {
    throw new CommandError(
      "formula-string",
      "String values starting with '=' are not allowed via xlsx:set-cell-value; use xlsx:set-cell-formula instead"
    );
  }
}
