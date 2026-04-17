import {
  CommandError,
  NotImplementedError,
  type CommandHandler,
} from "@officeai/core";
import type { DocxSnapshot } from "../model/types.js";
import { addCommentHandler } from "./add-comment.js";
import { deleteRangeHandler } from "./delete-range.js";
import { formatRangeHandler } from "./format-range.js";
import { insertParagraphHandler } from "./insert-paragraph.js";
import { insertTextHandler } from "./insert-text.js";
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
  makeStub("docx:insert-table", "Tables are P1; preserved on roundtrip but not yet editable."),
  makeStub("docx:set-cell-content", "Tables are P1; cell content editing not yet implemented."),
  makeStub("docx:insert-image", "Image insertion is P1; existing images are preserved."),
  makeStub("docx:resolve-comment", "Comment resolution is P1; only adding comments is supported."),
  makeStub("docx:accept-change", "Tracked-change accept is P1; revisions are preserved on roundtrip."),
  makeStub("docx:reject-change", "Tracked-change reject is P1; revisions are preserved on roundtrip."),
];

export const allDocxHandlers: ReadonlyArray<CommandHandler<unknown, DocxSnapshot>> = [
  insertTextHandler as CommandHandler<unknown, DocxSnapshot>,
  deleteRangeHandler as CommandHandler<unknown, DocxSnapshot>,
  formatRangeHandler as CommandHandler<unknown, DocxSnapshot>,
  insertParagraphHandler as CommandHandler<unknown, DocxSnapshot>,
  setParagraphStyleHandler as CommandHandler<unknown, DocxSnapshot>,
  addCommentHandler as CommandHandler<unknown, DocxSnapshot>,
  ...stubs,
];

export const docxHandlersById: ReadonlyMap<string, CommandHandler<unknown, DocxSnapshot>> =
  new Map(allDocxHandlers.map((h) => [h.type, h]));
