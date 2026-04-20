import type { CommandHandler } from "@officeai/core";
import type { PdfSnapshot } from "../model/types.js";
import {
  addAnnotationHandler,
  removeAnnotationHandler,
  updateAnnotationHandler,
} from "./annotation-commands.js";
import { addBookmarkHandler } from "./bookmark-commands.js";
import {
  addCommentHandler,
  deleteCommentHandler,
  editCommentHandler,
  replyCommentHandler,
  resolveCommentHandler,
} from "./comment-commands.js";
import { setMetadataHandler } from "./metadata-commands.js";
import {
  deletePagesHandler,
  reorderPagesHandler,
  rotatePagesHandler,
  setPageRotationHandler,
} from "./page-commands.js";

export * from "./payloads.js";
export {
  rotatePagesHandler,
  setPageRotationHandler,
  reorderPagesHandler,
  deletePagesHandler,
} from "./page-commands.js";
export { setMetadataHandler } from "./metadata-commands.js";
export { addBookmarkHandler } from "./bookmark-commands.js";
export {
  addCommentHandler,
  deleteCommentHandler,
  editCommentHandler,
  replyCommentHandler,
  resolveCommentHandler,
} from "./comment-commands.js";
export {
  addAnnotationHandler,
  removeAnnotationHandler,
  updateAnnotationHandler,
} from "./annotation-commands.js";

export const allPdfHandlers: ReadonlyArray<CommandHandler<unknown, PdfSnapshot>> = [
  rotatePagesHandler,
  setPageRotationHandler,
  reorderPagesHandler,
  deletePagesHandler,
  setMetadataHandler,
  addBookmarkHandler,
  addCommentHandler,
  replyCommentHandler,
  editCommentHandler,
  resolveCommentHandler,
  deleteCommentHandler,
  addAnnotationHandler,
  updateAnnotationHandler,
  removeAnnotationHandler,
] as ReadonlyArray<CommandHandler<unknown, PdfSnapshot>>;
