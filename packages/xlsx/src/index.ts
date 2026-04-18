/**
 * @officeai/xlsx — XLSX editor entry point.
 *
 * Phase 5 surface:
 *   - parseXlsx / serializeXlsx (byte-preserving when no edits; surgical
 *     re-emission for dirty sheets + workbook on rename)
 *   - typed cell layer (Cell, MergedCell, A1 reference utilities)
 *   - 5 P0 command handlers (`xlsx:set-cell-value`,
 *     `xlsx:set-range-values`, `xlsx:merge-cells`,
 *     `xlsx:unmerge-cells`, `xlsx:rename-sheet`)
 *
 * Surfaces still to ship:
 *   - Phase 6 — XlsxAgent (DocumentAgent contract); add-sheet + add-comment
 *   - Phase 7 — formula engine + set-cell-formula + insert/delete row/col
 *               + cross-sheet rename rewrite + style-table modeling
 *   - Phase 9 — virtualized renderer
 */

export const XLSX_PACKAGE_VERSION = "0.1.0";

export {
  parseXlsx,
  resolveTargetPath,
  XlsxParseError,
  type ParseOptions,
  type XlsxParseErrorCode,
} from "./parser/index.js";

export { serializeXlsx, XlsxSerializeError, type XlsxSerializeErrorCode } from "./serializer/index.js";

export {
  emptyDirty,
  cellKey,
  parseCellKey,
  parseA1,
  parseRange,
  formatA1,
  formatRange,
  colToLetter,
  letterToCol,
  rangeArea,
  rangesOverlap,
  type Cell,
  type CellAddress,
  type Comment,
  type CellErrorCode,
  type CellErrorValue,
  type CellRange,
  type CellValue,
  type Formula,
  type MergedCell,
  type OpaquePart,
  type Sheet,
  type XlsxDirtyFlags,
  type XlsxSnapshot,
  type XlsxWorkbook,
} from "./model/index.js";

export { flattenCellXf, type EffectiveStyle } from "./model/style-mutate.js";
export {
  type StyleAlignment,
  type StyleBorder,
  type StyleBorderSide,
  type StyleColor,
  type StyleFill,
  type StyleFont,
  type StyleTable,
} from "./model/style-table.js";

export {
  allXlsxHandlers,
  xlsxHandlersById,
  setCellValueHandler,
  setCellFormulaHandler,
  setCellFormatHandler,
  setRangeValuesHandler,
  mergeCellsHandler,
  unmergeCellsHandler,
  renameSheetHandler,
  addSheetHandler,
  insertRowHandler,
  insertColumnHandler,
  deleteRowHandler,
  deleteColumnHandler,
  addCommentHandler,
  setColumnWidthHandler,
  setRowHeightHandler,
  type SetCellValuePayload,
  type SetCellFormulaPayload,
  type SetCellFormatPayload,
  type CellFormatPatch,
  type CellFormatBorderSide,
  type SetRangeValuesPayload,
  type MergeCellsPayload,
  type UnmergeCellsPayload,
  type RenameSheetPayload,
  type AddSheetPayload,
  type InsertRowPayload,
  type InsertColumnPayload,
  type DeleteRowPayload,
  type DeleteColumnPayload,
  type AddCommentPayload,
  type SetColumnWidthPayload,
  type SetRowHeightPayload,
} from "./commands/index.js";

export {
  XlsxAgent,
  diffXlsxSnapshots,
  type XlsxAgentOptions,
  type XlsxRangeRequest,
  type XlsxRangeSnapshot,
  type XlsxSearchSpec,
  type XlsxSearchResult,
} from "./agent/index.js";

export {
  listRegisteredFunctions,
  type FunctionCategory,
  type RegisteredFunctionInfo,
} from "./formula/registered-functions.js";

export {
  tokenizeForDisplay,
  assignRefColors,
  DEFAULT_REF_COLORS,
  type DisplayToken,
  type DisplayTokenKind,
  type RefTarget,
} from "./formula/highlight.js";
