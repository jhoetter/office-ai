import type { CommandHandler } from "@officeai/core";
import type { XlsxSnapshot } from "../model/types.js";
import { mergeCellsHandler } from "./merge-cells.js";
import { renameSheetHandler } from "./rename-sheet.js";
import { setCellValueHandler } from "./set-cell-value.js";
import { setRangeValuesHandler } from "./set-range-values.js";
import { unmergeCellsHandler } from "./unmerge-cells.js";

/**
 * Phase 5 P0 command handlers.
 *
 * Shipped (5/13):
 *   - xlsx:set-cell-value
 *   - xlsx:set-range-values
 *   - xlsx:merge-cells
 *   - xlsx:unmerge-cells
 *   - xlsx:rename-sheet (without cross-sheet formula rewriting; that lands in Phase 7)
 *
 * Deferred to later phases (documented in `docs/build-log/xlsx.md`):
 *   - xlsx:set-cell-formula     → Phase 7 (formula engine)
 *   - xlsx:set-cell-format      → Phase 7+ (style table modeling)
 *   - xlsx:insert-row           → Phase 7 (formula adjustment)
 *   - xlsx:insert-column        → Phase 7 (formula adjustment)
 *   - xlsx:delete-row           → Phase 7 (formula adjustment)
 *   - xlsx:delete-column        → Phase 7 (formula adjustment)
 *   - xlsx:add-sheet            → Phase 6 (workbook + content-types + rels rewrite)
 *   - xlsx:add-comment          → Phase 6 (comments XML emission)
 */
export const allXlsxHandlers: ReadonlyArray<CommandHandler<unknown, XlsxSnapshot>> = [
  setCellValueHandler as CommandHandler<unknown, XlsxSnapshot>,
  setRangeValuesHandler as CommandHandler<unknown, XlsxSnapshot>,
  mergeCellsHandler as CommandHandler<unknown, XlsxSnapshot>,
  unmergeCellsHandler as CommandHandler<unknown, XlsxSnapshot>,
  renameSheetHandler as CommandHandler<unknown, XlsxSnapshot>,
];

export const xlsxHandlersById: ReadonlyMap<string, CommandHandler<unknown, XlsxSnapshot>> = new Map(
  allXlsxHandlers.map((h) => [h.type, h])
);
