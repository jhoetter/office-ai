import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { DeleteSheetPayload } from "./payloads.js";

/**
 * `xlsx:delete-sheet` — drop a worksheet (and its part files) from
 * the workbook.
 *
 * Pipeline:
 *   1. Validate: sheet must exist; cannot delete the only worksheet
 *      (Excel requires at least one visible sheet).
 *   2. Splice the sheet out of `workbook.sheets` and re-derive the
 *      `index` of every shifted neighbour.
 *   3. Sync the parallel SheetJS book (`SheetNames` + `Sheets`).
 *   4. Mark dirty:
 *        - `workbook` → re-emit the `<sheets>` list in `xl/workbook.xml`
 *        - `rels` → drop the orphan worksheet relationship in
 *          `xl/_rels/workbook.xml.rels`
 *        - `contentTypes` → drop the orphan `<Override>` in
 *          `[Content_Types].xml`
 *        - `removedSheetParts` → instruct the serializer to drop the
 *          `xl/worksheets/sheetN.xml` part and its `_rels/` sidecar
 *          from the package
 *
 * The remaining sheets keep their original part paths (so the part
 * indices in the zip may be non-contiguous, e.g. sheet1.xml + sheet3.xml
 * after deleting the middle sheet). That's a valid OOXML layout — the
 * relationship from `workbook.xml` to each sheet flows through `r:id`,
 * not through filename ordering.
 */
export const deleteSheetHandler: CommandHandler<DeleteSheetPayload, XlsxSnapshot> = {
  type: "xlsx:delete-sheet",
  apply(snapshot, payload) {
    const sheet = snapshot.root.sheets.find((s) => s.name === payload.name);
    if (!sheet) {
      throw new CommandError(
        "unknown-sheet",
        `Sheet "${payload.name}" not found; available: ${snapshot.root.sheets.map((s) => s.name).join(", ")}`
      );
    }

    const worksheetCount = snapshot.root.sheets.filter((s) => s.kind === "worksheet").length;
    if (sheet.kind === "worksheet" && worksheetCount <= 1) {
      throw new CommandError(
        "invalid-position",
        `Cannot delete the only worksheet "${sheet.name}"; a workbook must contain at least one visible sheet`
      );
    }

    const removedIndex = sheet.index;
    const sheets: Sheet[] = snapshot.root.sheets
      .filter((s) => s.id !== sheet.id)
      .map((s, i) => (s.index === i ? s : { ...s, index: i }));

    syncSheetjsDelete(snapshot.root, sheet.name);

    const nextWorkbook: XlsxWorkbook = { ...snapshot.root, sheets };
    const next = evolveSnapshot(snapshot, nextWorkbook, {
      workbook: true,
      rels: true,
      contentTypes: true,
      removedSheetParts: [sheet.partPath],
    });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-deleted",
          nodeId: sheet.id,
          path: ["sheets", removedIndex],
          summary: `Deleted sheet '${sheet.name}' (was index ${removedIndex})`,
          meta: {
            name: sheet.name,
            sheetId: sheet.sheetId,
            partPath: sheet.partPath,
          },
        },
      ]),
    };
  },
};

function syncSheetjsDelete(workbook: XlsxWorkbook, name: string): void {
  const book = workbook.sheetjs;
  const idx = book.SheetNames.indexOf(name);
  if (idx >= 0) book.SheetNames.splice(idx, 1);
  if (book.Sheets[name] !== undefined) {
    delete book.Sheets[name];
  }
}
