import { CommandError, type CommandHandler } from "@officeai/core";
import { formatRange } from "../model/refs.js";
import type { Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { UnmergeCellsPayload } from "./payloads.js";
import { parseRangeRef, resolveSheet } from "./validation.js";

export const unmergeCellsHandler: CommandHandler<UnmergeCellsPayload, XlsxSnapshot> = {
  type: "xlsx:unmerge-cells",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const range = parseRangeRef(payload.range);

    const idx = sheet.merges.findIndex(
      (m) =>
        m.r1 === range.start.row &&
        m.c1 === range.start.col &&
        m.r2 === range.end.row &&
        m.c2 === range.end.col
    );
    if (idx === -1) {
      throw new CommandError(
        "merge-not-found",
        `No exact-match merge at ${payload.range}; current merges: ${sheet.merges.length === 0 ? "(none)" : sheet.merges.map((m) => `${m.r1 + 1}:${m.c1 + 1}-${m.r2 + 1}:${m.c2 + 1}`).join(", ")}`
      );
    }

    const merges = sheet.merges.slice();
    merges.splice(idx, 1);
    const nextSheet: Sheet = { ...sheet, merges };
    const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
    const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-deleted",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "merges", formatRange(range)],
          summary: `unmerge ${formatRange(range)}`,
        },
      ]),
    };
  },
};
