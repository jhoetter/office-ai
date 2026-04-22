export * from "./payloads.js";
export { allDocxHandlers, docxHandlersById } from "./registry.js";
export { insertTextHandler } from "./insert-text.js";
export { deleteRangeHandler } from "./delete-range.js";
export { formatRangeHandler } from "./format-range.js";
export { insertParagraphHandler } from "./insert-paragraph.js";
export { setParagraphStyleHandler } from "./set-paragraph-style.js";
export { addCommentHandler } from "./add-comment.js";
export { paragraphPlainText } from "./helpers.js";
export { resolveCommentHandler } from "./resolve-comment.js";
export { replyCommentHandler } from "./reply-comment.js";
export { deleteCommentHandler } from "./delete-comment.js";
export { setHeaderTextHandler } from "./set-header-text.js";
export { setFooterTextHandler } from "./set-footer-text.js";
export { setHeaderFooterBlocksHandler } from "./set-header-footer-blocks.js";
export { createHeaderFooterPartHandler } from "./create-header-footer-part.js";
export { insertHeaderFooterImageHandler } from "./insert-header-footer-image.js";
export { acceptChangeHandler } from "./accept-change.js";
export { acceptAllChangesHandler } from "./accept-all-changes.js";
export { rejectChangeHandler } from "./reject-change.js";
export { rejectAllChangesHandler } from "./reject-all-changes.js";
export { insertTableHandler } from "./insert-table.js";
export { setCellContentHandler } from "./set-cell-content.js";
export { insertRowHandler } from "./insert-row.js";
export { insertColumnHandler } from "./insert-column.js";
export { deleteRowHandler, type DeleteRowPayload } from "./delete-row.js";
export { deleteColumnHandler, type DeleteColumnPayload } from "./delete-column.js";
export { deleteTableHandler, type DeleteTablePayload } from "./delete-table.js";
export { insertImageHandler } from "./insert-image.js";
export { applyListFormatHandler } from "./apply-list-format.js";
export { setParagraphListHandler } from "./set-paragraph-list.js";
export { removeParagraphListHandler } from "./remove-paragraph-list.js";
export { insertHyperlinkHandler } from "./insert-hyperlink.js";
export { removeHyperlinkHandler } from "./remove-hyperlink.js";
export { setParagraphAlignmentHandler } from "./set-paragraph-alignment.js";
export { setParagraphIndentHandler } from "./set-paragraph-indent.js";
export { setParagraphSpacingHandler, type SetParagraphSpacingPayload } from "./set-paragraph-spacing.js";
export { insertPageNumberHandler } from "./insert-page-number.js";
export { setSectionDifferentFirstHandler } from "./set-section-different-first.js";
export { setProtectionHandler } from "./set-protection.js";
export {
  mergeCellsHorizontalHandler,
  setCellAlignmentHandler,
  setCellShadingHandler,
  setColumnWidthHandler,
  setRowHeightHandler,
  type MergeCellsHorizontalPayload,
  type SetCellAlignmentPayload,
  type SetCellShadingPayload,
  type SetColumnWidthPayload,
  type SetRowHeightPayload,
} from "./set-table-cell-properties.js";
export { deleteImageHandler } from "./delete-image.js";
export { setImagePropertiesHandler } from "./set-image-properties.js";
export { setPageSetupHandler } from "./set-page-setup.js";
export { insertSectionBreakHandler } from "./insert-section-break.js";
export { insertPageBreakHandler, type InsertPageBreakPayload } from "./insert-page-break.js";
export { insertTextTrackedHandler } from "./insert-text-tracked.js";
export { deleteRangeTrackedHandler } from "./delete-range-tracked.js";
export { insertChartHandler } from "./insert-chart.js";
export { setChartDataHandler, setChartTitleHandler, setChartTypeHandler } from "./chart-edits.js";
export { insertSpreadsheetHandler, updateSpreadsheetHandler } from "./insert-spreadsheet.js";
export {
  insertFootnoteHandler,
  setFootnoteBodyHandler,
  deleteFootnoteHandler,
  type InsertFootnotePayload,
  type SetFootnoteBodyPayload,
  type DeleteFootnotePayload,
} from "./footnote-commands.js";
export {
  insertBookmarkHandler,
  deleteBookmarkHandler,
  listBookmarks,
  type BookmarkAnchor,
} from "./bookmarks.js";
