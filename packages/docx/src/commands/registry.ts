import type { CommandHandler } from "@officeai/core";
import type { DocxSnapshot } from "../model/types.js";
import { acceptChangeHandler } from "./accept-change.js";
import { addCommentHandler } from "./add-comment.js";
import { deleteCommentHandler } from "./delete-comment.js";
import { deleteRangeHandler } from "./delete-range.js";
import { formatRangeHandler } from "./format-range.js";
import { insertColumnHandler } from "./insert-column.js";
import { insertHyperlinkHandler } from "./insert-hyperlink.js";
import { insertImageHandler } from "./insert-image.js";
import { insertPageNumberHandler } from "./insert-page-number.js";
import { insertParagraphHandler } from "./insert-paragraph.js";
import { insertRowHandler } from "./insert-row.js";
import { insertSectionBreakHandler } from "./insert-section-break.js";
import { insertTableHandler } from "./insert-table.js";
import { insertTextHandler } from "./insert-text.js";
import { rejectChangeHandler } from "./reject-change.js";
import { removeHyperlinkHandler } from "./remove-hyperlink.js";
import { removeParagraphListHandler } from "./remove-paragraph-list.js";
import { replyCommentHandler } from "./reply-comment.js";
import { resolveCommentHandler } from "./resolve-comment.js";
import { setCellContentHandler } from "./set-cell-content.js";
import { setFooterTextHandler } from "./set-footer-text.js";
import { setHeaderTextHandler } from "./set-header-text.js";
import { setParagraphAlignmentHandler } from "./set-paragraph-alignment.js";
import { setParagraphIndentHandler } from "./set-paragraph-indent.js";
import { setParagraphListHandler } from "./set-paragraph-list.js";
import { setParagraphSpacingHandler } from "./set-paragraph-spacing.js";
import { setParagraphStyleHandler } from "./set-paragraph-style.js";
import { setSectionDifferentFirstHandler } from "./set-section-different-first.js";

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
  acceptChangeHandler as CommandHandler<unknown, DocxSnapshot>,
  rejectChangeHandler as CommandHandler<unknown, DocxSnapshot>,
  insertTableHandler as CommandHandler<unknown, DocxSnapshot>,
  setCellContentHandler as CommandHandler<unknown, DocxSnapshot>,
  insertRowHandler as CommandHandler<unknown, DocxSnapshot>,
  insertColumnHandler as CommandHandler<unknown, DocxSnapshot>,
  insertImageHandler as CommandHandler<unknown, DocxSnapshot>,
  setParagraphListHandler as CommandHandler<unknown, DocxSnapshot>,
  removeParagraphListHandler as CommandHandler<unknown, DocxSnapshot>,
  insertHyperlinkHandler as CommandHandler<unknown, DocxSnapshot>,
  removeHyperlinkHandler as CommandHandler<unknown, DocxSnapshot>,
  setParagraphAlignmentHandler as CommandHandler<unknown, DocxSnapshot>,
  setParagraphIndentHandler as CommandHandler<unknown, DocxSnapshot>,
  setParagraphSpacingHandler as CommandHandler<unknown, DocxSnapshot>,
  insertPageNumberHandler as CommandHandler<unknown, DocxSnapshot>,
  setSectionDifferentFirstHandler as CommandHandler<unknown, DocxSnapshot>,
  insertSectionBreakHandler as CommandHandler<unknown, DocxSnapshot>,
];

export const docxHandlersById: ReadonlyMap<string, CommandHandler<unknown, DocxSnapshot>> = new Map(
  allDocxHandlers.map((h) => [h.type, h])
);
