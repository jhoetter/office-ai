import { CommandError, type CommandHandler } from "@officeai/core";
import type { AutoFilter, FilterColumn, Sheet, XlsxSnapshot } from "../model/types.js";
import { recomputeHiddenRows } from "./auto-filter-eval.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SetFilterColumnPayload } from "./payloads.js";
import { resolveSheet } from "./validation.js";

/**
 * `xlsx:set-filter-column` — set / replace a single column's
 * criterion on the active AutoFilter and recompute hidden rows.
 *
 * Errors with `invalid-state` when no AutoFilter is active.
 */
export const setFilterColumnHandler: CommandHandler<SetFilterColumnPayload, XlsxSnapshot> = {
  type: "xlsx:set-filter-column",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    if (!sheet.autoFilter) {
      throw new CommandError(
        "invalid-state",
        `sheet "${sheet.name}" has no active AutoFilter; call xlsx:set-auto-filter first`
      );
    }

    const span = sheet.autoFilter.range.c2 - sheet.autoFilter.range.c1;
    if (!Number.isInteger(payload.colId) || payload.colId < 0 || payload.colId > span) {
      throw new CommandError(
        "invalid-range",
        `colId ${payload.colId} out of range; AutoFilter spans 0..${span}`
      );
    }

    validateCriterion(payload.criterion);

    const columns = new Map<number, FilterColumn>(sheet.autoFilter.columns);
    columns.set(payload.colId, payload.criterion);
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
          field: "criterion",
          summary: `${sheet.name} col ${payload.colId} filter → ${payload.criterion.kind}; ${hiddenRows.size} hidden`,
        },
      ]),
    };
  },
};

function validateCriterion(c: FilterColumn): void {
  switch (c.kind) {
    case "values":
      return;
    case "custom":
      if (!c.op1) throw new CommandError("invalid-payload", "custom filter requires op1");
      return;
    case "top10":
      if (!Number.isFinite(c.n) || c.n < 1 || c.n > 500) {
        throw new CommandError("invalid-payload", `top10.n must be in 1..500; got ${c.n}`);
      }
      return;
    case "dynamic":
      return;
    case "color":
      if (typeof c.argb !== "string" || c.argb.length === 0) {
        throw new CommandError("invalid-payload", "color filter requires non-empty argb");
      }
      return;
  }
}
