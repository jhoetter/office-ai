import { CommandError, type CommandHandler } from "@officeai/core";
import { formatRange } from "../model/refs.js";
import type { AutoFilter, Sheet, XlsxSnapshot } from "../model/types.js";
import { recomputeHiddenRows } from "./auto-filter-eval.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SetAutoFilterPayload } from "./payloads.js";
import { parseRangeRef, resolveSheet } from "./validation.js";

/**
 * `xlsx:set-auto-filter` — install / replace / remove the AutoFilter
 * band on a sheet.
 *
 * Pass `range: null` to remove the filter entirely (and unhide every
 * filter-driven hidden row). Setting a fresh range clears any
 * pre-existing per-column criteria; the handler intentionally does
 * NOT carry criteria over because Excel re-anchors the filter when
 * the range moves.
 */
export const setAutoFilterHandler: CommandHandler<SetAutoFilterPayload, XlsxSnapshot> = {
  type: "xlsx:set-auto-filter",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);

    let nextAutoFilter: AutoFilter | undefined;
    if (payload.range === null) {
      nextAutoFilter = undefined;
    } else {
      const range = parseRangeRef(payload.range);
      if (range.start.row === range.end.row) {
        throw new CommandError(
          "invalid-range",
          `auto-filter range "${payload.range}" must contain at least one body row below the header`
        );
      }
      nextAutoFilter = {
        range: {
          r1: range.start.row,
          c1: range.start.col,
          r2: range.end.row,
          c2: range.end.col,
        },
        columns: new Map(),
      };
    }

    const hiddenRows = recomputeHiddenRows(sheet, snapshot.root.styles, nextAutoFilter);

    const beforeDesc = sheet.autoFilter
      ? formatRange({
          start: { row: sheet.autoFilter.range.r1, col: sheet.autoFilter.range.c1 },
          end: { row: sheet.autoFilter.range.r2, col: sheet.autoFilter.range.c2 },
        })
      : "none";
    const afterDesc = nextAutoFilter
      ? formatRange({
          start: { row: nextAutoFilter.range.r1, col: nextAutoFilter.range.c1 },
          end: { row: nextAutoFilter.range.r2, col: nextAutoFilter.range.c2 },
        })
      : "none";

    if (beforeDesc === afterDesc && sheet.hiddenRows.size === hiddenRows.size) {
      const noop = evolveSnapshot(snapshot, snapshot.root, {});
      return { next: noop, diff: buildDiff(snapshot.revision, noop.revision, []) };
    }

    const nextSheet: Sheet = nextAutoFilter
      ? { ...sheet, autoFilter: nextAutoFilter, hiddenRows }
      : { ...stripAutoFilter(sheet), hiddenRows };

    const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
    const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "autoFilter"],
          field: "range",
          summary: `${sheet.name} autoFilter: ${beforeDesc} → ${afterDesc}`,
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
