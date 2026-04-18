import type { CommandHandler } from "@officeai/core";
import type { XlsxSnapshot } from "../model/types.js";
import { addCommentHandler } from "./add-comment.js";
import { addSheetHandler } from "./add-sheet.js";
import { deleteColumnHandler } from "./delete-column.js";
import { deleteRowHandler } from "./delete-row.js";
import { deleteSheetHandler } from "./delete-sheet.js";
import { insertColumnHandler } from "./insert-column.js";
import { insertRowHandler } from "./insert-row.js";
import { mergeCellsHandler } from "./merge-cells.js";
import { pasteRangeHandler } from "./paste-range.js";
import { renameSheetHandler } from "./rename-sheet.js";
import { setCellFormatHandler } from "./set-cell-format.js";
import { setCellFormulaHandler } from "./set-cell-formula.js";
import { setCellValueHandler } from "./set-cell-value.js";
import { setColumnWidthHandler } from "./set-column-width.js";
import { setRangeValuesHandler } from "./set-range-values.js";
import { setRowHeightHandler } from "./set-row-height.js";
import { unmergeCellsHandler } from "./unmerge-cells.js";

/**
 * P0 command handlers wired into the bus (13/13).
 *
 * Order matches the spec's §1–§13 sequence in
 * `spec/xlsx/agent-commands.md`:
 *   §1  xlsx:set-cell-value      (Phase 5)
 *   §2  xlsx:set-cell-formula    (Phase 7f; full recalc + cached writebacks)
 *   §3  xlsx:set-range-values    (Phase 5)
 *   §4  xlsx:set-cell-format     (Phase 7g; typed style table + content-hash dedupe)
 *   §5  xlsx:insert-row          (Phase 7i; cell shift + formula rewrite + recalc)
 *   §6  xlsx:insert-column       (Phase 7i)
 *   §7  xlsx:delete-row          (Phase 7i; #REF! casualty surfacing)
 *   §8  xlsx:delete-column       (Phase 7i)
 *   §9  xlsx:merge-cells         (Phase 5)
 *   §10 xlsx:unmerge-cells       (Phase 5)
 *   §11 xlsx:add-sheet           (Phase 7h; workbook + content-types + rels rewrite)
 *   §12 xlsx:rename-sheet        (Phase 5; cross-sheet formula rewriting deferred to 7+)
 *   §13 xlsx:add-comment         (Phase 7j; classic notes; threaded comments + VML deferred)
 */
export const allXlsxHandlers: ReadonlyArray<CommandHandler<unknown, XlsxSnapshot>> = [
  setCellValueHandler as CommandHandler<unknown, XlsxSnapshot>,
  setCellFormulaHandler as CommandHandler<unknown, XlsxSnapshot>,
  setRangeValuesHandler as CommandHandler<unknown, XlsxSnapshot>,
  setCellFormatHandler as CommandHandler<unknown, XlsxSnapshot>,
  insertRowHandler as CommandHandler<unknown, XlsxSnapshot>,
  insertColumnHandler as CommandHandler<unknown, XlsxSnapshot>,
  deleteRowHandler as CommandHandler<unknown, XlsxSnapshot>,
  deleteColumnHandler as CommandHandler<unknown, XlsxSnapshot>,
  mergeCellsHandler as CommandHandler<unknown, XlsxSnapshot>,
  unmergeCellsHandler as CommandHandler<unknown, XlsxSnapshot>,
  addSheetHandler as CommandHandler<unknown, XlsxSnapshot>,
  renameSheetHandler as CommandHandler<unknown, XlsxSnapshot>,
  addCommentHandler as CommandHandler<unknown, XlsxSnapshot>,
  setColumnWidthHandler as CommandHandler<unknown, XlsxSnapshot>,
  setRowHeightHandler as CommandHandler<unknown, XlsxSnapshot>,
  deleteSheetHandler as CommandHandler<unknown, XlsxSnapshot>,
  pasteRangeHandler as CommandHandler<unknown, XlsxSnapshot>,
];

export const xlsxHandlersById: ReadonlyMap<string, CommandHandler<unknown, XlsxSnapshot>> = new Map(
  allXlsxHandlers.map((h) => [h.type, h])
);
