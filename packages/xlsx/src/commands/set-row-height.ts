import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SetRowHeightPayload } from "./payloads.js";
import { resolveSheet } from "./validation.js";

const MIN_HEIGHT = 12;
const MAX_HEIGHT = 1024;

/**
 * `xlsx:set-row-height` — runtime UI affordance for the web grid
 * (P11g). Stores a per-row height override in CSS pixels on
 * `Sheet.rowHeights`. The OOXML `<row ht=…>` attribute stays opaque
 * in P0; the web Grid uses these overrides to render variable
 * row heights and the user can drag the row-resize handles to
 * update them. Pass `height: null` to reset to the Grid's default.
 */
export const setRowHeightHandler: CommandHandler<SetRowHeightPayload, XlsxSnapshot> = {
  type: "xlsx:set-row-height",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    if (!Number.isInteger(payload.row) || payload.row < 1 || payload.row > 1_048_576) {
      throw new CommandError(
        "invalid-range",
        `row must be a 1-based integer in 1..1048576; got ${payload.row}`
      );
    }
    if (payload.height !== null) {
      if (!Number.isFinite(payload.height) || payload.height < MIN_HEIGHT || payload.height > MAX_HEIGHT) {
        throw new CommandError(
          "invalid-range",
          `height must be in [${MIN_HEIGHT}..${MAX_HEIGHT}] px or null; got ${payload.height}`
        );
      }
    }

    const rowIndex0 = payload.row - 1;
    const before = sheet.rowHeights.get(rowIndex0);
    const next = payload.height === null ? undefined : Math.round(payload.height);
    if (before === next) {
      const noopNext = evolveSnapshot(snapshot, snapshot.root, {});
      return { next: noopNext, diff: buildDiff(snapshot.revision, noopNext.revision, []) };
    }

    const rowHeights = new Map(sheet.rowHeights);
    if (next === undefined) rowHeights.delete(rowIndex0);
    else rowHeights.set(rowIndex0, next);

    const nextSheet: Sheet = { ...sheet, rowHeights };
    const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
    const nextSnap = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });
    return {
      next: nextSnap,
      diff: buildDiff(snapshot.revision, nextSnap.revision, [
        {
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "rowHeights", String(payload.row)],
          field: "height",
          summary: `${sheet.name}!row${payload.row} height: ${before ?? "default"} → ${next ?? "default"}`,
          meta: { before: before ?? null, after: next ?? null },
        },
      ]),
    };
  },
};
