import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { RenameSheetPayload } from "./payloads.js";
import { assertUniqueSheetName, validateSheetName } from "./validation.js";

/**
 * Phase 5 rename: updates the typed sheet name + the SheetJS-side
 * worksheet key + the workbook's `SheetNames` array. Also marks the
 * workbook part dirty so the workbook XML's `<sheet name="...">`
 * attribute will be re-emitted by future phases.
 *
 * Cross-sheet formula reference rewriting (per `EC-R4`) lands in
 * Phase 7 alongside the formula engine; for now any formulas that
 * reference the renamed sheet by its OLD name will silently break on
 * recalc. Documented in `docs/build-log/xlsx.md`.
 */
export const renameSheetHandler: CommandHandler<RenameSheetPayload, XlsxSnapshot> = {
  type: "xlsx:rename-sheet",
  apply(snapshot, payload) {
    const sheet = snapshot.root.sheets.find((s) => s.name === payload.name);
    if (!sheet) {
      throw new CommandError(
        "unknown-sheet",
        `Sheet "${payload.name}" not found; available: ${snapshot.root.sheets.map((s) => s.name).join(", ")}`
      );
    }

    if (sheet.name === payload.newName) {
      const next = evolveSnapshot(snapshot, snapshot.root, {});
      return {
        next,
        diff: buildDiff(snapshot.revision, next.revision, [
          {
            kind: "node-updated",
            nodeId: sheet.id,
            path: ["sheets", sheet.index, "name"],
            field: "name",
            summary: `rename ${sheet.name} → ${payload.newName} (no-op)`,
          },
        ]),
      };
    }

    validateSheetName(payload.newName);
    assertUniqueSheetName(snapshot.root, payload.newName, sheet.index);

    const renamed: Sheet = { ...sheet, name: payload.newName };
    const sheets = snapshot.root.sheets.slice();
    sheets[sheet.index] = renamed;

    syncSheetjsRename(snapshot.root, sheet.name, payload.newName);

    const nextWorkbook: XlsxWorkbook = { ...snapshot.root, sheets };
    const next = evolveSnapshot(snapshot, nextWorkbook, {
      workbook: true,
      sheets: [sheet.partPath],
    });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "name"],
          field: "name",
          summary: `rename ${sheet.name} → ${payload.newName}`,
        },
      ]),
    };
  },
};

function syncSheetjsRename(workbook: XlsxWorkbook, oldName: string, newName: string): void {
  const book = workbook.sheetjs;
  const idx = book.SheetNames.indexOf(oldName);
  if (idx >= 0) book.SheetNames[idx] = newName;
  if (book.Sheets[oldName] !== undefined) {
    book.Sheets[newName] = book.Sheets[oldName];
    delete book.Sheets[oldName];
  }
}
