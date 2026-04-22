import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SetRowVisibilityPayload } from "./payloads.js";
import { resolveSheet } from "./validation.js";

/**
 * `xlsx:set-row-visibility` — show / hide a single row by adding or
 * removing it from `Sheet.hiddenRows`. The serializer's
 * `injectHiddenRows` paints `hidden="1"` on the matching `<row>`
 * element so the round-trip stays Excel-clean — we don't touch row
 * heights (Excel preserves the user's last height across hide/unhide).
 */
export const setRowVisibilityHandler: CommandHandler<SetRowVisibilityPayload, XlsxSnapshot> = {
  type: "xlsx:set-row-visibility",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    if (!Number.isInteger(payload.row) || payload.row < 1 || payload.row > 1_048_576) {
      throw new CommandError(
        "invalid-range",
        `row must be a 1-based integer in 1..1048576; got ${payload.row}`
      );
    }
    const rowIndex0 = payload.row - 1;
    const wasHidden = sheet.hiddenRows.has(rowIndex0);
    if (wasHidden === payload.hidden) {
      const noopNext = evolveSnapshot(snapshot, snapshot.root, {});
      return { next: noopNext, diff: buildDiff(snapshot.revision, noopNext.revision, []) };
    }
    const hiddenRows = new Set(sheet.hiddenRows);
    if (payload.hidden) hiddenRows.add(rowIndex0);
    else hiddenRows.delete(rowIndex0);
    const nextSheet: Sheet = { ...sheet, hiddenRows };
    const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
    const nextSnap = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });
    return {
      next: nextSnap,
      diff: buildDiff(snapshot.revision, nextSnap.revision, [
        {
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "hiddenRows", String(payload.row)],
          field: "hidden",
          summary: `${sheet.name}!row${payload.row} hidden: ${wasHidden} → ${payload.hidden}`,
          meta: { before: wasHidden, after: payload.hidden },
        },
      ]),
    };
  },
};
