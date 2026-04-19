/**
 * Client-side XLSX export helpers — CSV, TSV and JSON projections of
 * the typed snapshot. PDF/HTML go through the server convert route
 * (LibreOffice) instead; these helpers only need to cover the data
 * formats where round-trip layout doesn't matter.
 */

import JSZip from "jszip";
import type { Cell, Sheet, XlsxSnapshot } from "@officeai/xlsx";

const CSV_DELIMITER = ",";
const TSV_DELIMITER = "\t";

export interface SheetBounds {
  readonly rows: number;
  readonly cols: number;
}

/**
 * Convert a single sheet to delimiter-separated values. Empty cells
 * become empty strings; the row count tracks the highest row that
 * holds data so we don't emit gigabytes of trailing blank lines.
 */
export function sheetToCsv(
  sheet: Sheet,
  delimiter: string = CSV_DELIMITER
): string {
  const bounds = computeSheetBounds(sheet);
  if (bounds.rows === 0 || bounds.cols === 0) return "";
  const lines: string[] = [];
  for (let r = 0; r < bounds.rows; r += 1) {
    const cells: string[] = [];
    for (let c = 0; c < bounds.cols; c += 1) {
      const cell = sheet.cells.get(`${r}:${c}`);
      cells.push(escapeField(cellToString(cell), delimiter));
    }
    lines.push(cells.join(delimiter));
  }
  return lines.join("\r\n") + "\r\n";
}

export function sheetToTsv(sheet: Sheet): string {
  return sheetToCsv(sheet, TSV_DELIMITER);
}

/**
 * Bundle one CSV per sheet into a zip. Sheet names are sanitized for
 * safe filenames (forbidden characters replaced with `_`).
 */
export async function workbookToCsvZip(
  snapshot: XlsxSnapshot,
  delimiter: string = CSV_DELIMITER
): Promise<Blob> {
  const zip = new JSZip();
  const usedNames = new Set<string>();
  for (const sheet of snapshot.root.sheets) {
    if (sheet.kind !== "worksheet") continue;
    const baseName = safeFilename(sheet.name) || `Sheet${sheet.index + 1}`;
    let name = `${baseName}.csv`;
    let i = 2;
    while (usedNames.has(name)) {
      name = `${baseName}-${i}.csv`;
      i += 1;
    }
    usedNames.add(name);
    zip.file(name, sheetToCsv(sheet, delimiter));
  }
  return await zip.generateAsync({ type: "blob", mimeType: "application/zip" });
}

/**
 * Workbook → JSON. Each sheet becomes an array of row objects keyed
 * by the column letter (`A`, `B`, …) so the output stays usable when
 * the sheet doesn't have an explicit header row. Cells are typed
 * (numbers stay numbers, errors become `{ "#REF!": true }`-style
 * sentinels).
 */
export function workbookToJson(snapshot: XlsxSnapshot): string {
  const out: Record<string, ReadonlyArray<Record<string, unknown>>> = {};
  for (const sheet of snapshot.root.sheets) {
    if (sheet.kind !== "worksheet") continue;
    const bounds = computeSheetBounds(sheet);
    const rows: Record<string, unknown>[] = [];
    for (let r = 0; r < bounds.rows; r += 1) {
      const row: Record<string, unknown> = {};
      let nonEmpty = false;
      for (let c = 0; c < bounds.cols; c += 1) {
        const cell = sheet.cells.get(`${r}:${c}`);
        if (!cell) continue;
        const key = colToLetter(c);
        row[key] = cellToJson(cell);
        nonEmpty = true;
      }
      if (nonEmpty) rows.push(row);
    }
    out[sheet.name] = rows;
  }
  return JSON.stringify(out, null, 2);
}

export function computeSheetBounds(sheet: Sheet): SheetBounds {
  let maxRow = -1;
  let maxCol = -1;
  for (const cell of sheet.cells.values()) {
    if (cell.row > maxRow) maxRow = cell.row;
    if (cell.col > maxCol) maxCol = cell.col;
  }
  if (maxRow < 0 || maxCol < 0) return { rows: 0, cols: 0 };
  return { rows: maxRow + 1, cols: maxCol + 1 };
}

/* ── helpers ──────────────────────────────────────────────────────── */

function cellToString(cell: Cell | undefined): string {
  if (!cell) return "";
  return cellValueToString(cell.value);
}

function cellValueToString(value: Cell["value"]): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number") {
    return Number.isFinite(value) ? String(value) : "";
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "object" && "kind" in value && value.kind === "error") {
    return value.code;
  }
  return "";
}

function cellToJson(cell: Cell): unknown {
  const v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === "object" && "kind" in v && v.kind === "error") {
    return { error: v.code };
  }
  return v;
}

function escapeField(field: string, delimiter: string): string {
  if (field.length === 0) return "";
  const needsQuoting =
    field.includes(delimiter) ||
    field.includes('"') ||
    field.includes("\n") ||
    field.includes("\r");
  if (!needsQuoting) return field;
  return `"${field.replace(/"/g, '""')}"`;
}

function safeFilename(name: string): string {
  return name.replace(/[\u0000-\u001f\\/:*?"<>|]/g, "_").trim();
}

/** Excel-style column letters: 0 → A, 25 → Z, 26 → AA, … */
function colToLetter(col: number): string {
  let n = col + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}
