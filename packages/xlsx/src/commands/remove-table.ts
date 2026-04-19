import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { RemoveTablePayload } from "./payloads.js";
import { resolveSheet } from "./validation.js";

/**
 * `xlsx:remove-table` — strip an Excel Table from a sheet. Cells stay
 * put; this only removes the table object + its filter.
 *
 * Excel parity: "Table → Convert to Range" leaves cell values and
 * styles intact and removes the AutoFilter buttons. We mirror that
 * by clearing the autoFilter only when it matches the table's range
 * exactly — preserving any user-applied filter that might span more.
 */
export const removeTableHandler: CommandHandler<RemoveTablePayload, XlsxSnapshot> = {
  type: "xlsx:remove-table",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const tableIdx = sheet.tables.findIndex((t) => t.tableId === payload.tableId);
    if (tableIdx === -1) {
      throw new CommandError(
        "unknown-table",
        `Sheet "${sheet.name}" has no table with id ${payload.tableId}`
      );
    }
    const removed = sheet.tables[tableIdx];
    const tables = sheet.tables.slice();
    tables.splice(tableIdx, 1);

    const filterMatches =
      sheet.autoFilter !== undefined &&
      removed.autoFilterRange !== undefined &&
      formatAutoFilterRange(sheet.autoFilter) === removed.autoFilterRange;

    const nextSheet: Sheet = filterMatches
      ? { ...stripAutoFilter(sheet), tables, hiddenRows: new Set() }
      : { ...sheet, tables };

    const nextWorkbook: XlsxWorkbook = replaceSheet(snapshot.root, nextSheet);
    const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-deleted",
          nodeId: removed.id,
          path: ["sheets", sheet.index, "tables", tableIdx],
          summary: `Remove table ${removed.displayName} from ${sheet.name}`,
        },
      ]),
    };
  },
};

function stripAutoFilter(sheet: Sheet): Sheet {
  const { autoFilter: _autoFilter, ...rest } = sheet;
  void _autoFilter;
  return rest as Sheet;
}

function formatAutoFilterRange(filter: {
  range: { r1: number; c1: number; r2: number; c2: number };
}): string {
  const { r1, c1, r2, c2 } = filter.range;
  return `${colLetter(c1)}${r1 + 1}:${colLetter(c2)}${r2 + 1}`;
}

function colLetter(col: number): string {
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
