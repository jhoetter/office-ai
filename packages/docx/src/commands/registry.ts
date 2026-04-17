import { CommandError, NotImplementedError, type CommandHandler } from "@officeai/core";
import type { DocxSnapshot } from "../model/types.js";
import { acceptChangeHandler } from "./accept-change.js";
import { addCommentHandler } from "./add-comment.js";
import { deleteCommentHandler } from "./delete-comment.js";
import { deleteRangeHandler } from "./delete-range.js";
import { formatRangeHandler } from "./format-range.js";
import { insertColumnHandler } from "./insert-column.js";
import { insertParagraphHandler } from "./insert-paragraph.js";
import { insertRowHandler } from "./insert-row.js";
import { insertTableHandler } from "./insert-table.js";
import { insertTextHandler } from "./insert-text.js";
import { rejectChangeHandler } from "./reject-change.js";
import { replyCommentHandler } from "./reply-comment.js";
import { resolveCommentHandler } from "./resolve-comment.js";
import { setCellContentHandler } from "./set-cell-content.js";
import { setFooterTextHandler } from "./set-footer-text.js";
import { setHeaderTextHandler } from "./set-header-text.js";
import { setParagraphStyleHandler } from "./set-paragraph-style.js";

function makeStub<TPayload>(type: string, reason: string): CommandHandler<TPayload, DocxSnapshot> {
  return {
    type,
    apply() {
      throw new NotImplementedError(type, { reason });
    },
  };
}

void CommandError;

const stubs: ReadonlyArray<CommandHandler<unknown, DocxSnapshot>> = [
  makeStub("docx:insert-image", "Image insertion is P1; existing images are preserved."),
];

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
  ...stubs,
];

export const docxHandlersById: ReadonlyMap<string, CommandHandler<unknown, DocxSnapshot>> = new Map(
  allDocxHandlers.map((h) => [h.type, h])
);
