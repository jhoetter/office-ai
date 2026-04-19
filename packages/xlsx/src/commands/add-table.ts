import { CommandError, type CommandHandler } from "@officeai/core";
import { cellKey, colToLetter } from "../model/refs.js";
import type { AutoFilter, Cell, Sheet, TableDef, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { recomputeHiddenRows } from "./auto-filter-eval.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { AddTablePayload } from "./payloads.js";
import { parseRangeRef, resolveSheet } from "./validation.js";

const TABLE_NAME_RE = /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/;

/**
 * `xlsx:add-table` — promote a range to an Excel Table (ListObject).
 *
 * Excel parity:
 *   - Installs an AutoFilter at the table range so filter buttons
 *     show up immediately (matches Ctrl+T's default).
 *   - Reads or synthesises column display names so the typed model
 *     can render them in the column-banding overlay.
 *   - Mints a workbook-unique table name (`TableN` by default) and
 *     a workbook-unique numeric `tableId`. Both are required by
 *     OOXML — `displayName` is what shows in formulas.
 *
 * Round-trip caveat (C14 v1):
 *   - Existing tables in the source file round-trip byte-clean via
 *     the `opaqueParts` catch-all (the `xl/tables/` prefix is not
 *     in `MODELED_PREFIXES`). Tables added via this command live in
 *     the in-memory model and apply visually + via AutoFilter, but
 *     re-emitting a brand-new `xl/tables/tableN.xml` part on save
 *     ships in a follow-up. The user's data + filter survive saves;
 *     only the typed Table-ness is dropped on a save+reopen until
 *     then.
 */
export const addTableHandler: CommandHandler<AddTablePayload, XlsxSnapshot> = {
  type: "xlsx:add-table",
  apply(snapshot, payload, ctx) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const range = parseRangeRef(payload.range);

    const r1 = Math.min(range.start.row, range.end.row);
    const r2 = Math.max(range.start.row, range.end.row);
    const c1 = Math.min(range.start.col, range.end.col);
    const c2 = Math.max(range.start.col, range.end.col);

    const hasHeaders = payload.hasHeaders !== false;
    if (hasHeaders && r1 === r2) {
      throw new CommandError(
        "invalid-range",
        `add-table range "${payload.range}" must contain at least one body row below the header`
      );
    }

    const overlap = sheet.tables.find((t) => rangesOverlap(t.range, payload.range));
    if (overlap) {
      throw new CommandError(
        "invalid-range",
        `range "${payload.range}" overlaps existing table "${overlap.displayName}" at ${overlap.range}`
      );
    }

    const desiredName = payload.name ?? mintTableName(snapshot.root);
    if (!TABLE_NAME_RE.test(desiredName)) {
      throw new CommandError(
        "invalid-name",
        `table name "${desiredName}" is invalid; must match Excel's name rules`
      );
    }
    if (tableNameExists(snapshot.root, desiredName)) {
      throw new CommandError(
        "duplicate-name",
        `table name "${desiredName}" is already in use; table names must be unique workbook-wide`
      );
    }

    const tableId = mintTableId(snapshot.root);
    const partPath = mintTablePartPath(snapshot);
    const relId = mintTableRelId(sheet);

    const columnNames = readColumnNames(sheet, r1, c1, c2, hasHeaders);
    const a1Range = formatA1Range(r1, c1, r2, c2);

    const newTable: TableDef = {
      id: ctx.mintNodeId(),
      tableId: String(tableId),
      name: desiredName,
      displayName: desiredName,
      range: a1Range,
      headerRowCount: hasHeaders ? 1 : 0,
      totalsRowCount: 0,
      columnNames,
      autoFilterRange: a1Range,
      partPath,
      relId,
    };

    const nextAutoFilter: AutoFilter | undefined = hasHeaders
      ? { range: { r1, c1, r2, c2 }, columns: new Map() }
      : sheet.autoFilter;
    const hiddenRows = nextAutoFilter
      ? recomputeHiddenRows(sheet, snapshot.root.styles, nextAutoFilter)
      : sheet.hiddenRows;

    const nextSheet: Sheet = {
      ...sheet,
      tables: [...sheet.tables, newTable],
      ...(nextAutoFilter ? { autoFilter: nextAutoFilter, hiddenRows } : {}),
    };

    const nextWorkbook: XlsxWorkbook = replaceSheet(snapshot.root, nextSheet);
    const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-inserted",
          nodeId: newTable.id,
          path: ["sheets", sheet.index, "tables", sheet.tables.length],
          summary: `Add table ${desiredName} on ${sheet.name}!${a1Range}`,
          meta: { tableId: newTable.tableId, displayName: desiredName, range: a1Range },
        },
      ]),
    };
  },
};

function rangesOverlap(a: string, b: string): boolean {
  try {
    const A = parseRangeRef(a);
    const B = parseRangeRef(b);
    const ar1 = Math.min(A.start.row, A.end.row);
    const ar2 = Math.max(A.start.row, A.end.row);
    const ac1 = Math.min(A.start.col, A.end.col);
    const ac2 = Math.max(A.start.col, A.end.col);
    const br1 = Math.min(B.start.row, B.end.row);
    const br2 = Math.max(B.start.row, B.end.row);
    const bc1 = Math.min(B.start.col, B.end.col);
    const bc2 = Math.max(B.start.col, B.end.col);
    return !(ar2 < br1 || br2 < ar1 || ac2 < bc1 || bc2 < ac1);
  } catch {
    return false;
  }
}

function readColumnNames(
  sheet: Sheet,
  r1: number,
  c1: number,
  c2: number,
  hasHeaders: boolean
): ReadonlyArray<string> {
  const out: string[] = [];
  for (let c = c1; c <= c2; c++) {
    const offset = c - c1 + 1;
    if (!hasHeaders) {
      out.push(`Column${offset}`);
      continue;
    }
    const cell: Cell | undefined = sheet.cells.get(cellKey(r1, c));
    const label = formatHeader(cell);
    out.push(label.length > 0 ? label : `Column${offset}`);
  }
  return out;
}

function formatHeader(cell: Cell | undefined): string {
  if (!cell) return "";
  const v = cell.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v;
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  if (typeof v === "object" && "kind" in v && v.kind === "error") return "";
  return "";
}

function formatA1Range(r1: number, c1: number, r2: number, c2: number): string {
  return `${colToLetter(c1)}${r1 + 1}:${colToLetter(c2)}${r2 + 1}`;
}

function tableNameExists(workbook: XlsxWorkbook, name: string): boolean {
  const lower = name.toLowerCase();
  for (const s of workbook.sheets) {
    for (const t of s.tables) {
      if (t.displayName.toLowerCase() === lower || t.name.toLowerCase() === lower) return true;
    }
  }
  return false;
}

function mintTableName(workbook: XlsxWorkbook): string {
  for (let i = 1; i < 10_000; i++) {
    const candidate = `Table${i}`;
    if (!tableNameExists(workbook, candidate)) return candidate;
  }
  throw new CommandError("internal", "Could not mint a unique table name");
}

function mintTableId(workbook: XlsxWorkbook): number {
  let max = 0;
  for (const s of workbook.sheets) {
    for (const t of s.tables) {
      const n = Number.parseInt(t.tableId, 10);
      if (Number.isFinite(n) && n > max) max = n;
    }
  }
  return max + 1;
}

function mintTablePartPath(snapshot: XlsxSnapshot): string {
  const taken = new Set<string>();
  for (const s of snapshot.root.sheets) {
    for (const t of s.tables) taken.add(t.partPath);
  }
  for (const path of snapshot.container.parts.keys()) {
    if (path.startsWith("xl/tables/table") && path.endsWith(".xml")) taken.add(path);
  }
  for (let i = 1; i < 10_000; i++) {
    const candidate = `xl/tables/table${i}.xml`;
    if (!taken.has(candidate)) return candidate;
  }
  throw new CommandError("internal", "Could not mint a unique table part path");
}

function mintTableRelId(sheet: Sheet): string {
  const taken = new Set<string>();
  for (const t of sheet.tables) taken.add(t.relId);
  for (let i = 1; i < 10_000; i++) {
    const candidate = `rId${100 + i}`;
    if (!taken.has(candidate)) return candidate;
  }
  throw new CommandError("internal", "Could not mint a unique table rel id");
}
