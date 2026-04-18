import { CommandError, type CommandHandler } from "@officeai/core";
import { colToLetter } from "../model/refs.js";
import type { Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SetColumnWidthPayload } from "./payloads.js";
import { resolveSheet } from "./validation.js";

const MIN_WIDTH = 16;
const MAX_WIDTH = 4096;

/**
 * `xlsx:set-column-width` — runtime UI affordance for the web grid
 * (P11g). Stores a per-column width override in CSS pixels on
 * `Sheet.columnWidths`. The OOXML `<cols>` band stays opaque in P0;
 * the web Grid uses these overrides to render variable-geometry
 * columns and the user can drag the column-resize handles to update
 * them. Pass `width: null` to reset to the Grid's default.
 */
export const setColumnWidthHandler: CommandHandler<SetColumnWidthPayload, XlsxSnapshot> = {
  type: "xlsx:set-column-width",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    if (!Number.isInteger(payload.column) || payload.column < 1 || payload.column > 16384) {
      throw new CommandError(
        "invalid-range",
        `column must be a 1-based integer in 1..16384; got ${payload.column}`
      );
    }
    if (payload.width !== null) {
      if (!Number.isFinite(payload.width) || payload.width < MIN_WIDTH || payload.width > MAX_WIDTH) {
        throw new CommandError(
          "invalid-range",
          `width must be in [${MIN_WIDTH}..${MAX_WIDTH}] px or null; got ${payload.width}`
        );
      }
    }

    const colIndex0 = payload.column - 1;
    const before = sheet.columnWidths.get(colIndex0);
    const next = payload.width === null ? undefined : Math.round(payload.width);
    if (before === next) {
      const noopNext = evolveSnapshot(snapshot, snapshot.root, {});
      return { next: noopNext, diff: buildDiff(snapshot.revision, noopNext.revision, []) };
    }

    const columnWidths = new Map(sheet.columnWidths);
    if (next === undefined) columnWidths.delete(colIndex0);
    else columnWidths.set(colIndex0, next);

    const nextSheet: Sheet = { ...sheet, columnWidths };
    const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
    const nextSnap = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });
    return {
      next: nextSnap,
      diff: buildDiff(snapshot.revision, nextSnap.revision, [
        {
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "columnWidths", colToLetter(colIndex0)],
          field: "width",
          summary: `${sheet.name}!${colToLetter(colIndex0)} width: ${before ?? "default"} → ${next ?? "default"}`,
          meta: { before: before ?? null, after: next ?? null },
        },
      ]),
    };
  },
};
