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
  "docx:delete-row",
  "docx:delete-column",
  "docx:delete-table",
  "docx:insert-image",
  "docx:resolve-comment",
  "docx:reply-comment",
  "docx:delete-comment",
  "docx:accept-change",
  "docx:reject-change",
  "docx:set-header-text",
  "docx:set-footer-text",
  "docx:set-header-footer-blocks",
  "docx:create-header-footer-part",
  "docx:insert-header-footer-image",
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
  "docx:insert-chart",
  "docx:set-chart-data",
  "docx:set-chart-title",
  "docx:set-chart-type",
  "docx:insert-spreadsheet",
  "docx:update-spreadsheet",
  "docx:insert-footnote",
  "docx:set-footnote-body",
  "docx:delete-footnote",
  "docx:set-protection",
  "docx:set-cell-shading",
  "docx:set-cell-alignment",
  "docx:set-row-height",
  "docx:set-column-width",
  "docx:merge-cells-horizontal",
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

/* ── Chart authoring ─────────────────────────────────────────────────────── */

export type DocxChartKind = "bar" | "line" | "pie" | "area";

export interface ChartSeriesInput {
  /** Optional series legend label, displayed by Office's legend slot. */
  readonly name?: string;
  /** Numeric series data, one entry per category. */
  readonly values: ReadonlyArray<number>;
}

/**
 * Insert a typed chart at the given paragraph position. The handler:
 *
 *   1. Creates a new `word/charts/chartN.xml` part and a paired
 *      `word/embeddings/Microsoft_Excel_WorksheetN.xlsx` workbook so
 *      Office's "Edit Data" UI works on round-trip.
 *   2. Adds an `image`-style relationship from `word/document.xml.rels`
 *      to the chart part, plus a `package` relationship from the chart
 *      part to its embedded workbook (created by the serializer).
 *   3. Splices a typed `ChartDrawing` leaf into the targeted paragraph
 *      via the same run-splitting rules used by `docx:insert-image`.
 */
export interface InsertChartPayload {
  at: DocxPosition;
  chartType: DocxChartKind;
  categories: ReadonlyArray<string>;
  series: ReadonlyArray<ChartSeriesInput>;
  /** Optional chart title rendered above the plot area. */
  title?: string;
  /** Display width in **pixels** (96 DPI). Defaults to 480. */
  width?: number;
  /** Display height in **pixels**. Defaults to 320. */
  height?: number;
  /** Optional `<wp:docPr name>`. Defaults to `"Chart {docPrId}"`. */
  name?: string;
  /** Optional alt text for accessibility. */
  altText?: string;
}

/** Replace categories + series of an existing typed chart. */
export interface SetChartDataPayload {
  /** Part path of the chart, e.g. `"word/charts/chart1.xml"`. */
  chartPartPath: string;
  categories: ReadonlyArray<string>;
  series: ReadonlyArray<ChartSeriesInput>;
}

/** Set or clear the title of an existing typed chart. */
export interface SetChartTitlePayload {
  chartPartPath: string;
  /** `null` removes the title; a string sets it. */
  title: string | null;
}

/** Switch the active plot type of an existing typed chart. */
export interface SetChartTypePayload {
  chartPartPath: string;
  chartType: DocxChartKind;
}

/**
 * Insert an OLE-embedded Excel spreadsheet at the given position.
 *
 * The handler writes a new `word/embeddings/oleObjectN.xlsx` package
 * (built from the supplied 2D grid via `buildEmbeddedXlsx`) and a
 * paired `word/media/imageN.png` preview image (rendered from the
 * grid via `gridToPng`). It then registers an `oleObject` rel + an
 * `image` rel in `word/document.xml.rels` and splices a typed
 * `EmbeddedSpreadsheet` leaf into the targeted paragraph so Office
 * shows the preview until the user double-clicks to activate Excel.
 */
export interface InsertSpreadsheetPayload {
  at: DocxPosition;
  /**
   * 2D grid of cell values. Numbers become numeric `<v>` cells;
   * non-numeric values become inline strings. The first row is the
   * header band of the preview.
   */
  data: ReadonlyArray<ReadonlyArray<string | number | null | undefined>>;
  /** Optional worksheet name. Defaults to `"Sheet1"`. */
  sheetName?: string;
  /** Optional `<wp:docPr name>` analog. Defaults to `"Spreadsheet"`. */
  name?: string;
}

/**
 * Replace the embedded `.xlsx` bytes for an existing OLE spreadsheet
 * leaf. Used by the editor's "double-click → edit in transient
 * XlsxAgent → save" flow to push the modified workbook back into the
 * host document. The preview image is regenerated from `previewGrid`
 * when supplied; otherwise the previous preview is left in place.
 */
export interface UpdateSpreadsheetPayload {
  /** Part path of the embedded xlsx, e.g. `"word/embeddings/oleObject1.xlsx"`. */
  embeddingPartPath: string;
  /** Replacement xlsx package bytes. */
  bytes: Uint8Array;
  /** Optional new preview grid (re-renders the cached PNG). */
  previewGrid?: ReadonlyArray<ReadonlyArray<string | number | null | undefined>>;
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
 * Replace the entire `body` of a header or footer part. Used by
 * the Word-style in-place authoring path so a multi-paragraph
 * header (with page-number fields, inline images, etc.) commits
 * back to the model without flattening to a single text run the
 * way `set-header-text` / `set-footer-text` do.
 *
 * `partPath` doubles as the part id (e.g. `"word/header1.xml"`).
 * `body` must be non-empty — pass a single empty paragraph to
 * clear the part's content.
 */
export interface SetHeaderFooterBlocksPayload {
  partPath: string;
  body: ReadonlyArray<BlockNode>;
}

/**
 * Mint a fresh, empty header or footer part on the document and
 * attach it to the trailing implicit section as a `default` slot.
 *
 * Used by the in-place "double-click an empty header/footer zone"
 * affordance: the zone has no part backing it (the `body` zone is
 * Word-faithfully blank when no header/footer is referenced from
 * the section), so the editor first dispatches this command to
 * materialise a part, then focuses the freshly-rendered zone so
 * the caret lands and the user can start typing.
 *
 * Steps performed:
 * 1. Mint a fresh `word/header${N}.xml` (or `footer${N}.xml`) path
 *    that doesn't collide with any existing H/F part.
 * 2. Mint a fresh `rId${M}` for `word/document.xml.rels` and add
 *    the relationship pointing at the new part.
 * 3. Append a `headerReference` / `footerReference` of type `default`
 *    to the trailing section's `headerRefs` / `footerRefs` and drop
 *    that section's `raw` so the serializer rebuilds `<w:sectPr>`.
 * 4. Append a typed `HeaderFooterPart` with one empty paragraph to
 *    `headersAndFooters` (so the renderer immediately picks it up
 *    on the next snapshot tick).
 * 5. Set dirty flags: `body` (sectPr changed), `relationships`
 *    (`word/document.xml`), `headersAndFooters` (the new part), and
 *    `contentTypes` (the package needs an `Override` for the new
 *    part's content type).
 *
 * The `target` field exists primarily so a future "Different first
 * page" / "Different odd & even" UI can mint `first` / `even`
 * parts; this round only ever mints `default`.
 *
 * No-op (returns the same snapshot with a synthetic same-revision
 * diff) when the trailing section already has a default part of
 * the requested slot — the caller can safely fire this on every
 * focus and still get idempotent behavior.
 */
export interface CreateHeaderFooterPartPayload {
  slot: "header" | "footer";
  target?: "default" | "first" | "even";
}

/**
 * Insert an inline image into a header or footer part. Mirrors
 * `docx:insert-image` but the relationship lands inside the H/F
 * part's own rels file (`word/_rels/headerN.xml.rels`) instead of
 * the body's rels file. When `paragraphIndex` is omitted the image
 * is appended in a fresh paragraph at the end of the part.
 */
export interface InsertHeaderFooterImagePayload {
  /** Part path of the target H/F part, e.g. `"word/header1.xml"`. */
  partPath: string;
  /**
   * Optional 0-based paragraph index inside the part's body. When
   * present, the image is appended as a fresh run on that
   * paragraph (preserving its existing inlines). When absent, a
   * new paragraph holding only the image is appended.
   */
  paragraphIndex?: number;
  /** Raw bytes of the image. */
  data: Uint8Array | ArrayBuffer;
  mimeType: string;
  /** Display width in CSS pixels (96 DPI). */
  width: number;
  /** Display height in CSS pixels. */
  height: number;
  /** Optional alt text — populates `<wp:docPr descr>`. */
  altText?: string;
  /** Optional `<wp:docPr name>`. Defaults to `"Picture {docPrId}"`. */
  name?: string;
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

/* ── Footnotes (F1) ──────────────────────────────────────────────────────── */

export type {
  InsertFootnotePayload,
  SetFootnoteBodyPayload,
  DeleteFootnotePayload,
} from "./footnote-commands.js";

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

/**
 * `docx:set-protection` — toggle Word's "Restrict Editing" pane (the
 * `<w:documentProtection>` element on `word/settings.xml`). Mirrors
 * Word's Review tab → Restrict Editing → Yes, Start Enforcing
 * Protection workflow.
 *
 * No password hashing is performed by the handler. Pass a precomputed
 * `passwordHash` plus the matching `algorithmName` / `saltValue` /
 * `spinCount` to populate the OOXML hash attributes; otherwise the
 * element is written without a password (matching Word's "no password
 * required" option). Use `enabled: false` to drop the element
 * entirely.
 *
 * `edit` controls what the protection allows when enabled:
 *   • `readOnly`         — entire document is read-only
 *   • `comments`         — only comments may be authored
 *   • `trackedChanges`   — every edit is force-tracked
 *   • `forms`            — only form fields may be edited
 *
 * Spec: ECMA-376 Part 1, §17.15.1.29 (documentProtection).
 */
export interface SetProtectionPayload {
  readonly enabled: boolean;
  readonly edit?: "readOnly" | "comments" | "trackedChanges" | "forms";
  readonly enforce?: boolean;
  readonly formatting?: boolean;
  readonly algorithmName?: string;
  readonly passwordHash?: string;
  readonly saltValue?: string;
  readonly hashValue?: string;
  readonly spinCount?: number;
}
