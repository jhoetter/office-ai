import * as XLSX from "xlsx";
import { cellKey } from "../model/refs.js";
import type { Cell, CellErrorCode, MergedCell, Sheet } from "../model/types.js";

/**
 * Project a typed `Sheet`'s cells + merges back onto the SheetJS
 * `WorkSheet` it was originally derived from. Mutates `ws` in place.
 *
 * Phase 5 only handles literal value mutations + merge changes. The
 * formula path is wired up but does NOT recompute `v` — Phase 7's
 * formula engine fills that in. Style indices on existing cells are
 * preserved by reading them off the SheetJS cell and re-attaching.
 */
export function syncSheetToSheetJS(sheet: Sheet, ws: XLSX.WorkSheet): void {
  const dense = ws as unknown as MutableDenseSheet;

  let maxRow = -1;
  let maxCol = -1;
  for (const cell of sheet.cells.values()) {
    if (cell.row > maxRow) maxRow = cell.row;
    if (cell.col > maxCol) maxCol = cell.col;
  }

  const newData: Array<Array<SheetJSCell | undefined>> = [];
  const oldData = dense["!data"] ?? [];
  for (let r = 0; r <= maxRow; r++) {
    const oldRow = oldData[r];
    const newRow: Array<SheetJSCell | undefined> = [];
    for (let c = 0; c <= maxCol; c++) {
      const oldCell = oldRow ? oldRow[c] : undefined;
      const typed = sheet.cells.get(cellKey(r, c));
      newRow[c] = typed ? typedCellToSheetJS(typed, oldCell) : undefined;
    }
    newData[r] = newRow;
  }
  dense["!data"] = newData;

  if (sheet.merges.length > 0) {
    dense["!merges"] = sheet.merges.map((m: MergedCell) => ({
      s: { r: m.r1, c: m.c1 },
      e: { r: m.r2, c: m.c2 },
    }));
  } else {
    delete dense["!merges"];
  }

  if (maxRow >= 0 && maxCol >= 0) {
    dense["!ref"] = `A1:${colToA(maxCol)}${maxRow + 1}`;
  } else if (oldData.length === 0) {
    dense["!ref"] = "A1";
  }
}

function colToA(col: number): string {
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

interface SheetJSCell {
  t: string;
  v?: unknown;
  f?: string;
  s?: unknown;
  z?: string;
  w?: string;
}

interface MutableDenseSheet {
  "!data"?: Array<Array<SheetJSCell | undefined>>;
  "!merges"?: Array<{ s: { r: number; c: number }; e: { r: number; c: number } }>;
  "!ref"?: string;
}

const ERR_NUM_BY_CODE: Record<CellErrorCode, number> = {
  "#NULL!": 0x00,
  "#DIV/0!": 0x07,
  "#VALUE!": 0x0f,
  "#REF!": 0x17,
  "#NAME?": 0x1d,
  "#NUM!": 0x24,
  "#N/A": 0x2a,
  "#GETTING_DATA": 0x2b,
  "#SPILL!": 0x0f,
};

function typedCellToSheetJS(cell: Cell, prev: SheetJSCell | undefined): SheetJSCell | undefined {
  const carriedStyle = prev?.s !== undefined ? { s: prev.s } : {};
  const carriedFmt = prev?.z !== undefined ? { z: prev.z } : {};
  const v = cell.value;

  if (cell.formula) {
    return {
      t: typeof v === "number" ? "n" : typeof v === "string" ? "s" : typeof v === "boolean" ? "b" : "n",
      v: v ?? 0,
      f: cell.formula.text,
      ...carriedStyle,
      ...carriedFmt,
    };
  }

  if (typeof v === "number") {
    if (Number.isNaN(v)) return { t: "e", v: ERR_NUM_BY_CODE["#NUM!"], ...carriedStyle, ...carriedFmt };
    return { t: "n", v, ...carriedStyle, ...carriedFmt };
  }
  if (typeof v === "string") {
    return { t: "s", v, ...carriedStyle, ...carriedFmt };
  }
  if (typeof v === "boolean") {
    return { t: "b", v, ...carriedStyle, ...carriedFmt };
  }
  if (v === null) {
    if (prev?.s !== undefined) return { t: "z", ...carriedStyle, ...carriedFmt };
    return undefined;
  }
  if (typeof v === "object" && (v as { kind?: string }).kind === "error") {
    const code = (v as { code: CellErrorCode }).code;
    return { t: "e", v: ERR_NUM_BY_CODE[code] ?? ERR_NUM_BY_CODE["#VALUE!"], ...carriedStyle, ...carriedFmt };
  }
  return undefined;
}
