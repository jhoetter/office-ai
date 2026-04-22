import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SetColumnVisibilityPayload } from "./payloads.js";
import { resolveSheet } from "./validation.js";

/**
 * `xlsx:set-column-visibility` — show / hide a single column. Mutates
 * `Sheet.hiddenCols` AND patches `Sheet.colsXml` so the serializer
 * round-trips the change. Excel reads column hiddenness from
 * `<col hidden="1"/>` blocks inside `<cols>` (not from per-cell
 * shifts), so we update both the typed set and the opaque XML.
 *
 * Strategy:
 *   1. If a `<col min="X" max="X" .../>` (or a span containing X)
 *      already exists, set/clear its `hidden="1"` attribute.
 *   2. Else, append a fresh `<col min="X" max="X" width="9.140625"
 *      customWidth="1" hidden="1"/>` row.
 *   3. Always append `<cols>...</cols>` if `colsXml` was empty.
 */
export const setColumnVisibilityHandler: CommandHandler<SetColumnVisibilityPayload, XlsxSnapshot> = {
  type: "xlsx:set-column-visibility",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    if (!Number.isInteger(payload.column) || payload.column < 1 || payload.column > 16_384) {
      throw new CommandError(
        "invalid-range",
        `column must be a 1-based integer in 1..16384; got ${payload.column}`
      );
    }
    const colIndex0 = payload.column - 1;
    const wasHidden = sheet.hiddenCols.has(colIndex0);
    if (wasHidden === payload.hidden) {
      const noopNext = evolveSnapshot(snapshot, snapshot.root, {});
      return { next: noopNext, diff: buildDiff(snapshot.revision, noopNext.revision, []) };
    }
    const hiddenCols = new Set(sheet.hiddenCols);
    if (payload.hidden) hiddenCols.add(colIndex0);
    else hiddenCols.delete(colIndex0);

    const colsXml = patchColsXml(sheet.colsXml, payload.column, payload.hidden);
    const nextSheet: Sheet = {
      ...sheet,
      hiddenCols,
      ...(colsXml ? { colsXml } : {}),
    };
    const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
    const nextSnap = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });
    return {
      next: nextSnap,
      diff: buildDiff(snapshot.revision, nextSnap.revision, [
        {
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "hiddenCols", String(payload.column)],
          field: "hidden",
          summary: `${sheet.name}!col${payload.column} hidden: ${wasHidden} → ${payload.hidden}`,
          meta: { before: wasHidden, after: payload.hidden },
        },
      ]),
    };
  },
};

/**
 * Walk the `<cols>` band, parse each `<col min=A max=B .../>`, and
 * mutate the entry that covers `column1`. If the entry spans more
 * than the target column, it gets split (left/middle/right) so we
 * only flip the bit on the target. Returns a new XML string for
 * `Sheet.colsXml`.
 */
function patchColsXml(prev: string | undefined, column1: number, hidden: boolean): string {
  type Col = {
    min: number;
    max: number;
    hidden: boolean;
    extras: string;
  };
  const cols: Col[] = [];
  if (prev) {
    const colRegex = /<col\b([^/>]*)\/?>/g;
    let m: RegExpExecArray | null;
    while ((m = colRegex.exec(prev)) !== null) {
      const attrs = m[1];
      const minMatch = /\bmin="(\d+)"/.exec(attrs);
      const maxMatch = /\bmax="(\d+)"/.exec(attrs);
      const hiddenMatch = /\bhidden="(1|true)"/.exec(attrs);
      if (!minMatch || !maxMatch) continue;
      const min = Number.parseInt(minMatch[1], 10);
      const max = Number.parseInt(maxMatch[1], 10);
      const extras = attrs
        .replace(/\s*\bmin="\d+"/g, "")
        .replace(/\s*\bmax="\d+"/g, "")
        .replace(/\s*\bhidden="(?:1|true|0|false)"/g, "");
      cols.push({ min, max, hidden: !!hiddenMatch, extras });
    }
  }
  // Find a band that covers `column1`. Split if necessary so only
  // the target column flips.
  const idx = cols.findIndex((c) => column1 >= c.min && column1 <= c.max);
  if (idx >= 0) {
    const band = cols[idx];
    const splits: Col[] = [];
    if (band.min < column1) splits.push({ min: band.min, max: column1 - 1, hidden: band.hidden, extras: band.extras });
    splits.push({ min: column1, max: column1, hidden, extras: band.extras });
    if (band.max > column1) splits.push({ min: column1 + 1, max: band.max, hidden: band.hidden, extras: band.extras });
    cols.splice(idx, 1, ...splits);
  } else if (hidden) {
    cols.push({ min: column1, max: column1, hidden: true, extras: ' width="9.140625" customWidth="1"' });
  }
  // Sort + serialize.
  cols.sort((a, b) => a.min - b.min);
  const body = cols
    .map((c) => {
      const hAttr = c.hidden ? ' hidden="1"' : "";
      return `<col min="${c.min}" max="${c.max}"${c.extras}${hAttr}/>`;
    })
    .join("");
  return `<cols>${body}</cols>`;
}
