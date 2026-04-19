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
  "docx:set-paragraph-spacing",
  "docx:insert-page-number",
  "docx:set-section-different-first",
  "docx:insert-section-break",
  "docx:insert-page-break",
  "docx:insert-text-tracked",
  "docx:delete-range-tracked",
  "docx:set-page-setup",
  "docx:set-image-properties",
  "docx:delete-image",
  "docx:apply-list-format",
  "docx:accept-all-changes",
  "docx:reject-all-changes",
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

/** B8 — Accept every tracked change in the document. No payload fields. */
export type AcceptAllChangesPayload = Record<string, never>;

/** B8 — Reject every tracked change in the document. No payload fields. */
export type RejectAllChangesPayload = Record<string, never>;

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

/**
 * B7 — high-level "Bullet list" / "Numbered list" toolbar command.
 *
 * Unlike {@link SetParagraphListPayload} (which requires the caller
 * to know a concrete `numId` from `word/numbering.xml`), this command
 * takes the user-facing intent (`"bullet"` vs `"decimal"`), looks for
 * a compatible existing `<w:num>` instance and falls back to
 * auto-minting a fresh `NumberingDefinitions` entry — including
 * registering `word/numbering.xml` in the package's relationships
 * graph and `[Content_Types].xml` overrides on save. This is what
 * the Word-style toolbar buttons dispatch.
 */
export interface ApplyListFormatPayload {
  paragraphId: NodeId;
  format: "bullet" | "decimal";
  ilvl?: number;
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

/**
 * Inserts a `<w:fldSimple w:instr=" PAGE \\* MERGEFORMAT "/>` (or
 * `NUMPAGES`) into a header / footer paragraph at a flat-text byte
 * offset. Errors with `unknown-target` if the target paragraph does
 * not live inside a header or footer part — page-number fields in the
 * body are valid OOXML but are not yet a P3 surface.
 */
export interface InsertPageNumberPayload {
  /** Stable id of the target paragraph (must be inside a header/footer part). */
  paragraphId: NodeId;
  /** Byte offset inside the paragraph's flat-text. Clamped to [0, length]. */
  offset: number;
  /** Defaults to "PAGE". */
  field?: "PAGE" | "NUMPAGES";
}

/**
 * Toggle `<w:titlePg/>` on the section containing `paragraphIndex`.
 * Walks forward from `paragraphIndex` to find the next
 * {@link SectionBreak}; if none is found, falls back to the trailing
 * implicit section at the end of the body.
 */
export interface SetSectionDifferentFirstPayload {
  paragraphIndex: number;
  enabled: boolean;
}

/**
 * Insert a fresh {@link SectionBreak} block at `paragraphIndex`,
 * inheriting the next section's geometry (page size, margins,
 * header/footer refs). The new break carries `<w:type>` set to
 * `type`, defaulting to `"nextPage"`.
 */
export interface InsertSectionBreakPayload {
  paragraphIndex: number;
  type?: "nextPage" | "continuous" | "evenPage" | "oddPage";
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

/**
 * Tracked-changes ("Suggesting"-mode) variant of `docx:insert-text`.
 * The inserted run is wrapped in a `<w:ins>` revision wrapper so the
 * insertion shows up as a green underline in Word and Google Docs and
 * can later be `accept`ed (folded into the body) or `reject`ed (the
 * inserted text disappears).
 *
 * Author / date populate the wrapper's `<w:ins w:author w:date>`
 * attributes so the existing `TrackedChangesUI` ribbon can attribute
 * the suggestion. `revisionId` is optional; when omitted the handler
 * mints a synthetic `mint-{n}` id derived from the existing revisions
 * in the snapshot so it stays unique across the document.
 */
export interface InsertTextTrackedPayload {
  at: DocxPosition;
  text: string;
  author: string;
  /** ISO 8601 string. Defaults to "now" at handler time. */
  date?: string;
  /** Caller-controlled revision id; minted when absent. */
  revisionId?: string;
}

/**
 * Tracked-changes variant of `docx:delete-range`. Instead of removing
 * the targeted text, the runs (or run segments) covering the range
 * are wrapped in a `<w:del>` revision wrapper and their text leaves
 * are flipped to `isDelText: true` so the serializer emits
 * `<w:delText>` (Word's struck-through display).
 *
 * MVP scope: single-paragraph ranges only. Multi-paragraph deletions
 * fall back to a `not-implemented` `CommandError`; tracked deletes
 * across paragraph boundaries also need a `<w:p>` end-marker
 * revision (`<w:rPr><w:del/></w:rPr>` on the paragraph mark) which is
 * deferred to a follow-up.
 */
export interface DeleteRangeTrackedPayload {
  range: DocxSelection;
  author: string;
  date?: string;
  revisionId?: string;
}

/**
 * B3 — Page Setup. Updates any subset of `pgSz` / `pgMar` on the
 * section that owns `paragraphIndex`. Walks forward to the next
 * `<w:sectPr>`; falls back to the trailing implicit section.
 *
 * Margins / sizes are in TWIPS. Pass only the fields you want to
 * change; omitted fields preserve the current value.
 */
/**
 * B6 — Update display dimensions / accessibility metadata for an
 * inline image. Width/height are in CSS pixels (96 DPI). Pass a
 * field as `undefined` to leave it untouched, or `null` to clear an
 * optional value (e.g. `altText: null` removes the alt text).
 */
/**
 * B6 — Remove an inline image leaf by id. The owning run + paragraph
 * survive; only the image leaf is excised.
 */
export interface DeleteImagePayload {
  imageId: string;
}

export interface SetImagePropertiesPayload {
  imageId: string;
  widthPx?: number | null;
  heightPx?: number | null;
  altText?: string | null;
  name?: string;
}

export interface SetPageSetupPayload {
  paragraphIndex: number;
  pgSz?: {
    w?: number;
    h?: number;
    orient?: "portrait" | "landscape";
  };
  pgMar?: {
    top?: number;
    right?: number;
    bottom?: number;
    left?: number;
    header?: number;
    footer?: number;
    gutter?: number;
  };
}
