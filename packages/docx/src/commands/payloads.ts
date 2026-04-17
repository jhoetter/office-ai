import type { BlockNode, DocxPosition, DocxSelection } from "../model/types.js";

export const DOCX_COMMAND_TYPES = [
  "docx:insert-text",
  "docx:delete-range",
  "docx:format-range",
  "docx:insert-paragraph",
  "docx:set-paragraph-style",
  "docx:add-comment",
  "docx:insert-table",
  "docx:set-cell-content",
  "docx:insert-image",
  "docx:resolve-comment",
  "docx:reply-comment",
  "docx:delete-comment",
  "docx:accept-change",
  "docx:reject-change",
] as const;

export type DocxCommandType = (typeof DOCX_COMMAND_TYPES)[number];

export interface InsertTextPayload {
  at: DocxPosition;
  text: string;
}

export interface DeleteRangePayload {
  range: DocxSelection;
}

export interface TextFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontFamily?: string;
  fontSize?: number;
  color?: string;
  highlight?: string;
}

export interface FormatRangePayload {
  range: DocxSelection;
  format: TextFormat;
}

export interface InsertParagraphPayload {
  at: DocxPosition;
  style?: string;
}

export interface SetParagraphStylePayload {
  at: DocxPosition;
  style: string;
}

export interface AddCommentPayload {
  range: DocxSelection;
  text: string;
  author: string;
  initials?: string;
}

export interface InsertTablePayload {
  at: DocxPosition;
  rows: number;
  cols: number;
}

export interface SetCellContentPayload {
  tableId: string;
  row: number;
  col: number;
  content: BlockNode[];
}

export interface InsertImagePayload {
  at: DocxPosition;
  data: ArrayBuffer;
  mimeType: string;
  width: number;
  height: number;
}

export interface ResolveCommentPayload {
  commentId: string;
  /** Defaults to true. Pass `false` to re-open a previously resolved comment. */
  resolved?: boolean;
}

export interface ReplyCommentPayload {
  parentId: string;
  text: string;
  author: string;
  initials?: string;
}

export interface DeleteCommentPayload {
  commentId: string;
}

export interface AcceptChangePayload {
  revisionId: string;
}

export interface RejectChangePayload {
  revisionId: string;
}
