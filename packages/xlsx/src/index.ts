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
  type AutoFilter,
  type AutoFilterRange,
  type Cell,
  type CellAddress,
  type Comment,
  type CellErrorCode,
  type CellErrorValue,
  type CellRange,
  type CellValue,
  type CustomFilterOp,
  type DynamicFilterType,
  type FilterColumn,
  type FilterColumnColor,
  type FilterColumnCustom,
  type FilterColumnDynamic,
  type FilterColumnTop10,
  type FilterColumnValues,
  type Formula,
  type FreezePanes,
  type ConditionalFormat,
  type ConditionalFormatOverlay,
  type DataValidation,
  type DefinedName,
  type TableDef,
  type ChartKind,
  type ChartPalette,
  type SheetChart,
  type MergedCell,
  type OpaquePart,
  type PivotCachePart,
  type PivotTableLocation,
  type PivotTablePart,
  type Sheet,
  type XlsxDirtyFlags,
  type XlsxSnapshot,
  type XlsxWorkbook,
  type ImageAnchor,
  type ImageBlob,
  type ImageContentType,
  type SheetImage,
  EMU_PER_PX,
  pxToEmu,
  emuToPx,
  contentTypeForExtension,
  EXTENSION_BY_CONTENT_TYPE,
} from "./model/index.js";

export {
  evaluateConditionalFormats,
  type OverlayMap as ConditionalFormatOverlayMap,
} from "./conditional-format/evaluate.js";

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
  moveSheetHandler,
  setSheetStateHandler,
  addConditionalFormatHandler,
  removeConditionalFormatHandler,
  clearConditionalFormatsHandler,
  addDataValidationHandler,
  removeDataValidationHandler,
  clearDataValidationsHandler,
  addDefinedNameHandler,
  updateDefinedNameHandler,
  removeDefinedNameHandler,
  insertRowHandler,
  insertColumnHandler,
  deleteRowHandler,
  deleteColumnHandler,
  addCommentHandler,
  addImageHandler,
  moveImageHandler,
  resizeImageHandler,
  removeImageHandler,
  setColumnWidthHandler,
  setRowHeightHandler,
  deleteSheetHandler,
  pasteRangeHandler,
  setAutoFilterHandler,
  setFilterColumnHandler,
  clearFilterColumnHandler,
  freezePanesHandler,
  unfreezePanesHandler,
  addTableHandler,
  removeTableHandler,
  addChartHandler,
  removeChartHandler,
  moveChartHandler,
  resizeChartHandler,
  sortRangeHandler,
  removeDuplicatesHandler,
  recomputeHiddenRows,
  type SetAutoFilterPayload,
  type SetFilterColumnPayload,
  type ClearFilterColumnPayload,
  type FreezePanesPayload,
  type UnfreezePanesPayload,
  type SortRangePayload,
  type RemoveDuplicatesPayload,
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
  type MoveSheetPayload,
  type SetSheetStatePayload,
  type AddConditionalFormatPayload,
  type RemoveConditionalFormatPayload,
  type ClearConditionalFormatsPayload,
  type AddDataValidationPayload,
  type RemoveDataValidationPayload,
  type ClearDataValidationsPayload,
  type AddDefinedNamePayload,
  type UpdateDefinedNamePayload,
  type RemoveDefinedNamePayload,
  type InsertRowPayload,
  type InsertColumnPayload,
  type DeleteRowPayload,
  type DeleteColumnPayload,
  type AddCommentPayload,
  type AddImagePayload,
  type MoveImagePayload,
  type ResizeImagePayload,
  type RemoveImagePayload,
  type ReplyCommentPayload,
  type ResolveCommentPayload,
  type DeleteCommentPayload,
  type EditCommentPayload,
  type SetColumnWidthPayload,
  type SetRowHeightPayload,
  type DeleteSheetPayload,
  type PasteRangePayload,
  type FillRangePayload,
  type TextToColumnsPayload,
  type AddTablePayload,
  type RemoveTablePayload,
  type AddChartPayload,
  type RemoveChartPayload,
  type MoveChartPayload,
  type ResizeChartPayload,
} from "./commands/index.js";

export {
  extractClipboardSnapshot,
  snapshotToTsv,
  tsvToSnapshot,
  delimitedToSnapshot,
  sniffDelimiter,
  type ClipboardMerge,
  type XlsxClipboardCell,
  type XlsxClipboardSnapshot,
} from "./clipboard/snapshot.js";

export { parseExternalClipboard, parseHtmlTable, parseFingerprintHtml } from "./clipboard/external.js";

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

export { xlsxActions, xlsxActionsById, xlsxActionsByCommandType } from "./actions/index.js";

export {
  tokenizeForDisplay,
  assignRefColors,
  DEFAULT_REF_COLORS,
  type DisplayToken,
  type DisplayTokenKind,
  type RefTarget,
} from "./formula/highlight.js";

export {
  buildEmbeddedXlsx,
  buildChartGrid,
  type BuildEmbeddedXlsxOptions,
  type BuildEmbeddedXlsxResult,
  type ChartGridResult,
  type ChartGridSeries,
  type EmbeddedGrid,
  type EmbeddedGridCell,
} from "./embed/buildEmbeddedXlsx.js";

export { gridToPng, type GridToPngOptions, type GridToPngResult } from "./embed/gridToPng.js";
