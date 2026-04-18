import { CommandError, type CommandHandler } from "@officeai/core";
import type { AutoFilter, FilterColumn, Sheet, XlsxSnapshot } from "../model/types.js";
import { recomputeHiddenRows } from "./auto-filter-eval.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { ClearFilterColumnPayload } from "./payloads.js";
import { resolveSheet } from "./validation.js";

/**
 * `xlsx:clear-filter-column` — drop the criterion on a single column
 * of the active AutoFilter and recompute hidden rows. No-op if the
 * column has no criterion.
 */
export const clearFilterColumnHandler: CommandHandler<ClearFilterColumnPayload, XlsxSnapshot> = {
  type: "xlsx:clear-filter-column",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    if (!sheet.autoFilter) {
      throw new CommandError(
        "invalid-state",
        `sheet "${sheet.name}" has no active AutoFilter`
      );
    }
    if (!sheet.autoFilter.columns.has(payload.colId)) {
      const noop = evolveSnapshot(snapshot, snapshot.root, {});
      return { next: noop, diff: buildDiff(snapshot.revision, noop.revision, []) };
    }

    const columns = new Map<number, FilterColumn>(sheet.autoFilter.columns);
    columns.delete(payload.colId);
    const nextAutoFilter: AutoFilter = { ...sheet.autoFilter, columns };
    const hiddenRows = recomputeHiddenRows(sheet, snapshot.root.styles, nextAutoFilter);

    const nextSheet: Sheet = { ...sheet, autoFilter: nextAutoFilter, hiddenRows };
    const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
    const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "autoFilter", "columns", String(payload.colId)],
          field: "cleared",
          summary: `${sheet.name} col ${payload.colId} filter cleared; ${hiddenRows.size} hidden`,
        },
      ]),
    };
  },
};
