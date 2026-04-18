import type { BlockNode, DocxPosition, DocxSelection, TableProperties } from "../model/types.js";
import type { NodeId } from "@officeai/core";

export const DOCX_COMMAND_TYPES = [
  "docx:insert-text",
  "docx:delete-range",
  "docx:format-range",
  "docx:insert-paragraph",
  "docx:set-paragraph-style",
  "docx:add-comment",
  "docx:insert-table",
  "docx:set-cell-content",
  "docx:insert-row",
  "docx:insert-column",
  "docx:insert-image",
  "docx:resolve-comment",
  "docx:reply-comment",
  "docx:delete-comment",
  "docx:accept-change",
  "docx:reject-change",
  "docx:set-header-text",
  "docx:set-footer-text",
  "docx:set-paragraph-list",
  "docx:remove-paragraph-list",
  "docx:insert-hyperlink",
  "docx:remove-hyperlink",
  "docx:set-paragraph-alignment",
  "docx:set-paragraph-indent",
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
  /** Optional explicit column widths in twips. Length must equal `cols`. */
  columnWidths?: number[];
  /** Optional table-level properties applied verbatim. */
  properties?: Partial<TableProperties>;
}

export interface SetCellContentPayload {
  tableId: string;
  row: number;
  col: number;
  content: BlockNode[];
}

export interface InsertRowPayload {
  tableId: string;
  /** 0-based row index. `at === rows.length` appends. */
  at: number;
}

export interface InsertColumnPayload {
  tableId: string;
  /** 0-based column index. `at === grid.length` appends. */
  at: number;
  /** Optional column width in twips. Defaults to an equal-split. */
  width?: number;
}

export interface InsertImagePayload {
  /**
   * Position of the new image leaf. The image is inserted as a fresh run
   * inside the targeted paragraph: with `run` + `offset` set, the
   * existing run is split at `offset` and the image run is spliced
   * between the two halves; without them, the image becomes the first
   * run of the paragraph (matching `docx:insert-text` semantics).
   */
  at: DocxPosition;
  /** Raw bytes of the image. Encoded as `Uint8Array` or `ArrayBuffer`. */
  data: Uint8Array | ArrayBuffer;
  /** MIME type, e.g. `image/png`, `image/jpeg`, `image/gif`. */
  mimeType: string;
  /** Display width in **pixels** (96 DPI). Converted to EMUs internally. */
  width: number;
  /** Display height in pixels. */
  height: number;
  /** Optional alt text — populates `<wp:docPr descr>`. */
  altText?: string;
  /** Optional `<wp:docPr name>`. Defaults to `"Picture {docPrId}"`. */
  name?: string;
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

export interface SetHeaderTextPayload {
  /**
   * Stable id of the header part to mutate. Equals the OOXML part path,
   * e.g. `"word/header1.xml"`. Discoverable via
   * `snapshot.root.headersAndFooters[i].id`.
   */
  partId: string;
  /** 0-based index into the header part's `body` array. */
  paragraphIndex: number;
  /** New plain-text content for the targeted paragraph. */
  text: string;
}

export interface SetFooterTextPayload {
  partId: string;
  paragraphIndex: number;
  text: string;
}

export interface SetParagraphListPayload {
  /** Stable id of the target paragraph (body or table cell). */
  paragraphId: NodeId;
  /** Concrete `<w:num>` instance id from `word/numbering.xml`. */
  numId: number;
  /** 0-based level inside the abstract definition. */
  ilvl: number;
}

export interface RemoveParagraphListPayload {
  paragraphId: NodeId;
}

export interface InsertHyperlinkPayload {
  paragraphId: NodeId;
  /** Flat-text byte range inside the paragraph (`start < end`, both inclusive of the paragraph length). */
  range: { start: number; end: number };
  /** External URL; mints a fresh `external` rel. Mutually exclusive with `anchor`. */
  url?: string;
  /** Internal bookmark name. Mutually exclusive with `url`. */
  anchor?: string;
}

export interface RemoveHyperlinkPayload {
  hyperlinkId: NodeId;
}

export interface SetParagraphAlignmentPayload {
  /** Stable id of the paragraph (body or table cell). */
  paragraphId: NodeId;
  /**
   * `null` clears the alignment, falling back to the document/style
   * default (which Word normally renders as left-to-right left-aligned).
   */
  alignment: "left" | "center" | "right" | "justify" | null;
}

export interface SetParagraphIndentPayload {
  /** Stable id of the paragraph (body or table cell). */
  paragraphId: NodeId;
  /**
   * Signed delta in twips applied to the paragraph's `indentation.left`.
   * The handler clamps the result to the OOXML legal range
   * `[0, 31680]` twips (≈ 22 inches).
   *
   * Pass a positive delta to "increase indent", negative to "outdent".
   * Standard Word toolbar steps use ±360 twips (¼ inch).
   */
  deltaTwips: number;
}
