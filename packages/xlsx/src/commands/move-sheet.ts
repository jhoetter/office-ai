import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { MoveSheetPayload } from "./payloads.js";
import type * as XLSX from "@e965/xlsx";

/**
 * `xlsx:move-sheet` — reorder a worksheet within the workbook.
 *
 * Mirrors Excel's "Move or Copy Sheet…" + drag-to-reorder. The
 * destination is clamped to `[0, sheets.length-1]`. Cross-sheet
 * formula references continue to resolve because Excel anchors them
 * by sheet name, not position.
 *
 * Pipeline:
 *   1. Locate the sheet by name. Compute the clamped destination.
 *   2. Splice the sheet out of `workbook.sheets` and reinsert it at
 *      `to`. Re-derive the `index` of every shifted neighbour.
 *   3. Mirror the move in the SheetJS workbook (`SheetNames` array)
 *      so derived getters that still go through SheetJS see the new
 *      ordering.
 *   4. Mark `dirty.workbook = true`. The serializer's
 *      `rewriteWorkbookSheets` re-emits `<sheets>` from the new
 *      `workbook.sheets` order; every other byte of `xl/workbook.xml`
 *      is preserved by string-level splice.
 */
export const moveSheetHandler: CommandHandler<MoveSheetPayload, XlsxSnapshot> = {
  type: "xlsx:move-sheet",
  apply(snapshot, payload) {
    const sheets = snapshot.root.sheets;
    const from = sheets.findIndex((s) => s.name === payload.name);
    if (from === -1) {
      throw new CommandError(
        "unknown-sheet",
        `Sheet "${payload.name}" not found; available: ${sheets.map((s) => s.name).join(", ")}`
      );
    }
    if (!Number.isFinite(payload.to)) {
      throw new CommandError("invalid-position", `Destination ${payload.to} is not a finite number`);
    }
    const to = Math.max(0, Math.min(sheets.length - 1, Math.floor(payload.to)));
    if (from === to) {
      const next = evolveSnapshot(snapshot, snapshot.root, {});
      return {
        next,
        diff: buildDiff(snapshot.revision, next.revision, [
          {
            kind: "node-updated",
            nodeId: sheets[from].id,
            path: ["sheets", from],
            field: "index",
            summary: `move sheet ${payload.name} (no-op, already at ${to})`,
          },
        ]),
      };
    }

    const next = sheets.slice();
    const [moved] = next.splice(from, 1);
    next.splice(to, 0, moved);
    const reindexed: Sheet[] = next.map((s, i) => (s.index === i ? s : { ...s, index: i }));

    syncSheetjsMove(snapshot.root, payload.name, to);

    const nextWorkbook: XlsxWorkbook = { ...snapshot.root, sheets: reindexed };
    const evolved = evolveSnapshot(snapshot, nextWorkbook, { workbook: true });

    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, [
        {
          kind: "node-updated",
          nodeId: moved.id,
          path: ["sheets", to],
          field: "index",
          summary: `move sheet ${payload.name}: ${from} → ${to}`,
        },
      ]),
    };
  },
};

function syncSheetjsMove(workbook: XlsxWorkbook, name: string, to: number): void {
  const book = workbook.sheetjs as { SheetNames: string[]; Sheets: Record<string, XLSX.WorkSheet> };
  const idx = book.SheetNames.indexOf(name);
  if (idx === -1) return;
  const [n] = book.SheetNames.splice(idx, 1);
  const clamped = Math.max(0, Math.min(book.SheetNames.length, to));
  book.SheetNames.splice(clamped, 0, n);
}
