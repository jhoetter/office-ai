import type { CommandHandler } from "@officeai/core";
import type { DocxSnapshot } from "../model/types.js";
import { acceptAllChangesHandler } from "./accept-all-changes.js";
import { acceptChangeHandler } from "./accept-change.js";
import { addCommentHandler } from "./add-comment.js";
import { applyListFormatHandler } from "./apply-list-format.js";
import { setChartDataHandler, setChartTitleHandler, setChartTypeHandler } from "./chart-edits.js";
import { insertChartHandler } from "./insert-chart.js";
import { insertSpreadsheetHandler, updateSpreadsheetHandler } from "./insert-spreadsheet.js";
import { rejectAllChangesHandler } from "./reject-all-changes.js";
import { deleteCommentHandler } from "./delete-comment.js";
import { deleteRangeHandler } from "./delete-range.js";
import { deleteRangeTrackedHandler } from "./delete-range-tracked.js";
import { formatRangeHandler } from "./format-range.js";
import { deleteColumnHandler } from "./delete-column.js";
import { deleteRowHandler } from "./delete-row.js";
import { deleteTableHandler } from "./delete-table.js";
import { insertColumnHandler } from "./insert-column.js";
import { insertHyperlinkHandler } from "./insert-hyperlink.js";
import { insertImageHandler } from "./insert-image.js";
import { insertPageBreakHandler } from "./insert-page-break.js";
import { insertPageNumberHandler } from "./insert-page-number.js";
import { insertParagraphHandler } from "./insert-paragraph.js";
import { insertRowHandler } from "./insert-row.js";
import { insertSectionBreakHandler } from "./insert-section-break.js";
import { insertTableHandler } from "./insert-table.js";
import { insertTextHandler } from "./insert-text.js";
import { insertTextTrackedHandler } from "./insert-text-tracked.js";
import { rejectChangeHandler } from "./reject-change.js";
import { removeHyperlinkHandler } from "./remove-hyperlink.js";
import { removeParagraphListHandler } from "./remove-paragraph-list.js";
import { replyCommentHandler } from "./reply-comment.js";
import { resolveCommentHandler } from "./resolve-comment.js";
import { setCellContentHandler } from "./set-cell-content.js";
import { setFooterTextHandler } from "./set-footer-text.js";
import { setHeaderTextHandler } from "./set-header-text.js";
import { setHeaderFooterBlocksHandler } from "./set-header-footer-blocks.js";
import { createHeaderFooterPartHandler } from "./create-header-footer-part.js";
import { insertHeaderFooterImageHandler } from "./insert-header-footer-image.js";
import { setParagraphAlignmentHandler } from "./set-paragraph-alignment.js";
import { setParagraphIndentHandler } from "./set-paragraph-indent.js";
import { setParagraphListHandler } from "./set-paragraph-list.js";
import { setParagraphSpacingHandler } from "./set-paragraph-spacing.js";
import { setParagraphStyleHandler } from "./set-paragraph-style.js";
import { deleteImageHandler } from "./delete-image.js";
import { setImagePropertiesHandler } from "./set-image-properties.js";
import { setPageSetupHandler } from "./set-page-setup.js";
import { setProtectionHandler } from "./set-protection.js";
import { setSectionDifferentFirstHandler } from "./set-section-different-first.js";
import {
  mergeCellsHorizontalHandler,
  setCellAlignmentHandler,
  setCellShadingHandler,
  setColumnWidthHandler,
  setRowHeightHandler,
} from "./set-table-cell-properties.js";
import { deleteFootnoteHandler, insertFootnoteHandler, setFootnoteBodyHandler } from "./footnote-commands.js";

export const allDocxHandlers: ReadonlyArray<CommandHandler<unknown, DocxSnapshot>> = [
  insertTextHandler as CommandHandler<unknown, DocxSnapshot>,
  deleteRangeHandler as CommandHandler<unknown, DocxSnapshot>,
  formatRangeHandler as CommandHandler<unknown, DocxSnapshot>,
  insertParagraphHandler as CommandHandler<unknown, DocxSnapshot>,
  setParagraphStyleHandler as CommandHandler<unknown, DocxSnapshot>,
  addCommentHandler as CommandHandler<unknown, DocxSnapshot>,
  resolveCommentHandler as CommandHandler<unknown, DocxSnapshot>,
  replyCommentHandler as CommandHandler<unknown, DocxSnapshot>,
  deleteCommentHandler as CommandHandler<unknown, DocxSnapshot>,
  setHeaderTextHandler as CommandHandler<unknown, DocxSnapshot>,
  setFooterTextHandler as CommandHandler<unknown, DocxSnapshot>,
  setHeaderFooterBlocksHandler as CommandHandler<unknown, DocxSnapshot>,
  createHeaderFooterPartHandler as CommandHandler<unknown, DocxSnapshot>,
  insertHeaderFooterImageHandler as CommandHandler<unknown, DocxSnapshot>,
  acceptChangeHandler as CommandHandler<unknown, DocxSnapshot>,
  rejectChangeHandler as CommandHandler<unknown, DocxSnapshot>,
  acceptAllChangesHandler as CommandHandler<unknown, DocxSnapshot>,
  rejectAllChangesHandler as CommandHandler<unknown, DocxSnapshot>,
  insertTableHandler as CommandHandler<unknown, DocxSnapshot>,
  setCellContentHandler as CommandHandler<unknown, DocxSnapshot>,
  insertRowHandler as CommandHandler<unknown, DocxSnapshot>,
  insertColumnHandler as CommandHandler<unknown, DocxSnapshot>,
  deleteRowHandler as CommandHandler<unknown, DocxSnapshot>,
  deleteColumnHandler as CommandHandler<unknown, DocxSnapshot>,
  deleteTableHandler as CommandHandler<unknown, DocxSnapshot>,
  insertImageHandler as CommandHandler<unknown, DocxSnapshot>,
  setParagraphListHandler as CommandHandler<unknown, DocxSnapshot>,
  removeParagraphListHandler as CommandHandler<unknown, DocxSnapshot>,
  applyListFormatHandler as CommandHandler<unknown, DocxSnapshot>,
  insertHyperlinkHandler as CommandHandler<unknown, DocxSnapshot>,
  removeHyperlinkHandler as CommandHandler<unknown, DocxSnapshot>,
  setParagraphAlignmentHandler as CommandHandler<unknown, DocxSnapshot>,
  setParagraphIndentHandler as CommandHandler<unknown, DocxSnapshot>,
  setParagraphSpacingHandler as CommandHandler<unknown, DocxSnapshot>,
  insertPageNumberHandler as CommandHandler<unknown, DocxSnapshot>,
  setSectionDifferentFirstHandler as CommandHandler<unknown, DocxSnapshot>,
  setPageSetupHandler as CommandHandler<unknown, DocxSnapshot>,
  setImagePropertiesHandler as CommandHandler<unknown, DocxSnapshot>,
  deleteImageHandler as CommandHandler<unknown, DocxSnapshot>,
  insertSectionBreakHandler as CommandHandler<unknown, DocxSnapshot>,
  insertPageBreakHandler as CommandHandler<unknown, DocxSnapshot>,
  insertTextTrackedHandler as CommandHandler<unknown, DocxSnapshot>,
  deleteRangeTrackedHandler as CommandHandler<unknown, DocxSnapshot>,
  insertChartHandler as CommandHandler<unknown, DocxSnapshot>,
  setChartDataHandler as CommandHandler<unknown, DocxSnapshot>,
  setChartTitleHandler as CommandHandler<unknown, DocxSnapshot>,
  setChartTypeHandler as CommandHandler<unknown, DocxSnapshot>,
  insertSpreadsheetHandler as CommandHandler<unknown, DocxSnapshot>,
  updateSpreadsheetHandler as CommandHandler<unknown, DocxSnapshot>,
  insertFootnoteHandler as CommandHandler<unknown, DocxSnapshot>,
  setFootnoteBodyHandler as CommandHandler<unknown, DocxSnapshot>,
  deleteFootnoteHandler as CommandHandler<unknown, DocxSnapshot>,
  setProtectionHandler as CommandHandler<unknown, DocxSnapshot>,
  setCellShadingHandler as CommandHandler<unknown, DocxSnapshot>,
  setCellAlignmentHandler as CommandHandler<unknown, DocxSnapshot>,
  setRowHeightHandler as CommandHandler<unknown, DocxSnapshot>,
  setColumnWidthHandler as CommandHandler<unknown, DocxSnapshot>,
  mergeCellsHorizontalHandler as CommandHandler<unknown, DocxSnapshot>,
];

export const docxHandlersById: ReadonlyMap<string, CommandHandler<unknown, DocxSnapshot>> = new Map(
  allDocxHandlers.map((h) => [h.type, h])
);
