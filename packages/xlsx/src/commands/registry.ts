import type { CommandHandler } from "@officeai/core";
import type { XlsxSnapshot } from "../model/types.js";
import { addSheetHandler } from "./add-sheet.js";
import { deleteColumnHandler } from "./delete-column.js";
import { deleteRowHandler } from "./delete-row.js";
import { insertColumnHandler } from "./insert-column.js";
import { insertRowHandler } from "./insert-row.js";
import { mergeCellsHandler } from "./merge-cells.js";
import { renameSheetHandler } from "./rename-sheet.js";
import { setCellFormatHandler } from "./set-cell-format.js";
import { setCellFormulaHandler } from "./set-cell-formula.js";
import { setCellValueHandler } from "./set-cell-value.js";
import { setRangeValuesHandler } from "./set-range-values.js";
import { unmergeCellsHandler } from "./unmerge-cells.js";

/**
 * P0 command handlers wired into the bus.
 *
 * Shipped (12/13):
 *   - xlsx:set-cell-value      (Phase 5)
 *   - xlsx:set-range-values    (Phase 5)
 *   - xlsx:merge-cells         (Phase 5)
 *   - xlsx:unmerge-cells       (Phase 5)
 *   - xlsx:rename-sheet        (Phase 5; cross-sheet formula rewriting deferred to 7+)
 *   - xlsx:set-cell-formula    (Phase 7f; runs full recalc and writes downstream cached values)
 *   - xlsx:set-cell-format     (Phase 7g; typed style table + content-hash dedupe)
 *   - xlsx:add-sheet           (Phase 7h; workbook + content-types + rels rewrite)
 *   - xlsx:insert-row          (Phase 7i; cell shift + formula rewrite + recalc)
 *   - xlsx:insert-column       (Phase 7i)
 *   - xlsx:delete-row          (Phase 7i; #REF! casualty surfacing)
 *   - xlsx:delete-column       (Phase 7i)
 *
 * Deferred to later sub-phases (documented in `docs/build-log/xlsx.md`):
 *   - xlsx:add-comment          → Phase 7j (comments XML emission)
 */
export const allXlsxHandlers: ReadonlyArray<CommandHandler<unknown, XlsxSnapshot>> = [
  setCellValueHandler as CommandHandler<unknown, XlsxSnapshot>,
  setCellFormulaHandler as CommandHandler<unknown, XlsxSnapshot>,
  setRangeValuesHandler as CommandHandler<unknown, XlsxSnapshot>,
  setCellFormatHandler as CommandHandler<unknown, XlsxSnapshot>,
  mergeCellsHandler as CommandHandler<unknown, XlsxSnapshot>,
  unmergeCellsHandler as CommandHandler<unknown, XlsxSnapshot>,
  renameSheetHandler as CommandHandler<unknown, XlsxSnapshot>,
  addSheetHandler as CommandHandler<unknown, XlsxSnapshot>,
  insertRowHandler as CommandHandler<unknown, XlsxSnapshot>,
  insertColumnHandler as CommandHandler<unknown, XlsxSnapshot>,
  deleteRowHandler as CommandHandler<unknown, XlsxSnapshot>,
  deleteColumnHandler as CommandHandler<unknown, XlsxSnapshot>,
];

export const xlsxHandlersById: ReadonlyMap<string, CommandHandler<unknown, XlsxSnapshot>> = new Map(
  allXlsxHandlers.map((h) => [h.type, h])
);
