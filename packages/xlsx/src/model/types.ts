import type { DocumentSnapshot, NodeId, ooxml } from "@officeai/core";
import type { ImageBlob, SheetImage } from "./drawings.js";
import type { StyleTable } from "./style-table.js";

/**
 * XLSX in-memory model — Phase 5 surface.
 *
 * Phase 5 grows the Phase 4 thin model with a typed cell store + merge
 * regions per sheet, sufficient to author the value/structure-mutation
 * commands (`xlsx:set-cell-value`, `xlsx:set-range-values`,
 * `xlsx:add-sheet`, `xlsx:rename-sheet`, `xlsx:merge-cells`,
 * `xlsx:unmerge-cells`). Styles, formulas, comments, hyperlinks, and
 * conditional formats stay opaque for now and land in Phase 7+.
 *
 * The SheetJS `WorkBook` (`XlsxWorkbook.sheetjs`) remains the
 * authoritative cell-layer wire format; the typed cells here are
 * derived at parse time and projected back to SheetJS at serialize
 * time when a sheet is dirty.
 */

export interface XlsxSnapshot extends DocumentSnapshot<XlsxWorkbook> {
  readonly format: "xlsx";
  /** Internal: the OOXML container. Not serialized; not part of equality. */
  readonly container: ooxml.OoxmlContainer;
  /** Per-part dirty flags consulted by the serializer. */
  readonly dirty: XlsxDirtyFlags;
}

export interface XlsxDirtyFlags {
  /** `xl/workbook.xml` — sheet list, defined names, workbook properties. */
  workbook: boolean;
  /** `xl/sharedStrings.xml`. */
  sharedStrings: boolean;
  /** `xl/styles.xml`. */
  styles: boolean;
  /** `[Content_Types].xml`. */
  contentTypes: boolean;
  /** `_rels/.rels` and `xl/_rels/workbook.xml.rels`. */
  rels: boolean;
  /** Per-sheet dirty set — keyed by part path (`xl/worksheets/sheetN.xml`). */
  sheets: ReadonlySet<string>;
  /** Per-comments-part dirty set — keyed by part path (`xl/commentsN.xml`). */
  comments: ReadonlySet<string>;
  /** Per-threaded-comments dirty set — keyed by part path. */
  threadedComments: ReadonlySet<string>;
  /** Per-sheet rels parts (`xl/worksheets/_rels/sheetN.xml.rels`). */
  sheetRels: ReadonlySet<string>;
  /**
   * Sheet part paths the next serializer pass MUST drop from the
   * package: the `xl/worksheets/sheetN.xml` part, its `_rels/`
   * sidecar, the workbook-rels relationship that targeted it, and
   * the content-types `<Override>` that registered it. Driven by
   * `xlsx:delete-sheet`. Membership implies the sheet has already
   * been removed from `workbook.sheets`.
   */
  removedSheetParts: ReadonlySet<string>;
  /**
   * Sheet part paths whose drawing part needs to be re-emitted
   * (`xl/drawings/drawingN.xml` + its `_rels`). Driven by
   * `xlsx:add-image` / `move-image` / `resize-image` / `remove-image`.
   * Membership implies the sheet's `<drawing>` reference + content-types
   * override are managed alongside the dirty pass.
   */
  drawings: ReadonlySet<string>;
  /**
   * Media part paths (`xl/media/imageN.png`) the next serializer pass
   * must (re)write. Brand-new media is added; bytes for existing
   * media are overwritten on edit. Removed media ends up in
   * `removedMediaParts` instead.
   */
  media: ReadonlySet<string>;
  /**
   * Media part paths the serializer must drop entirely (no sheet
   * references them anymore). Mirrors `removedSheetParts`'s shape.
   */
  removedMediaParts: ReadonlySet<string>;
}

export interface XlsxWorkbook {
  readonly id: NodeId;
  /** Sheets in tab order. Order matters; `index` mirrors the array index. */
  readonly sheets: ReadonlyArray<Sheet>;
  /**
   * SHA-256 hex per zip-entry path. The byte-preservation oracle.
   * Same shape as `DocumentSnapshot.partHashes`; doubly stored on the
   * root because the serializer reads from the workbook.
   */
  readonly partHashes: Readonly<Record<string, string>>;
  /**
   * Opaque parts keyed by full zip path. Anything the model does not
   * type — charts, drawings, embeddings, VBA, custom XML, themes,
   * pivots, slicers, …. Round-trips byte-identical via the container
   * cache.
   */
  readonly opaqueParts: ReadonlyMap<string, OpaquePart>;
  /**
   * `WorkbookPr.date1904` flag. When `true`, all date serials are
   * relative to 1904-01-01 instead of 1900-01-00. We do not normalize.
   */
  readonly date1904: boolean;
  /**
   * Original namespace declarations + non-content root attributes from
   * `xl/workbook.xml`. Re-emitted verbatim when the workbook part is
   * dirty.
   */
  readonly workbookRootAttrs: Readonly<Record<string, string>>;
  /**
   * Typed style table parsed from `xl/styles.xml`. When the workbook
   * has no styles part the parser substitutes `defaultStyleTable()`
   * so commands can always intern through the same shape. Re-emitted
   * by the serializer when `dirty.styles` is set.
   */
  readonly styles: StyleTable;
  /**
   * Phase 4 escape hatch: the SheetJS `WorkBook` produced by
   * `XLSX.read(..., { dense: true, ... })` carrying the cell layer.
   * Phase 5 introduces a typed `cells: Map<string, Cell>` on each
   * `Sheet` and this field becomes derived. Consumers should NOT take
   * a hard dependency on its shape outside `parser/`, `serializer/`,
   * and the Phase 5 model upgrade path.
   */
  readonly sheetjs: import("xlsx").WorkBook;
  /**
   * Image media keyed by `xl/media/...` part path. Sheet-level image
   * placements reference into this map via `SheetImage.mediaRef`, so
   * identical bytes uploaded twice round-trip through a single media
   * part (deduped on `ImageBlob.hash`).
   */
  readonly images: ReadonlyMap<string, ImageBlob>;
}

export interface Sheet {
  readonly id: NodeId;
  /** Stable OOXML sheet id (`<sheet sheetId="…">`). String, not the index. */
  readonly sheetId: string;
  /** Display name as it appears on the tab. UTF-8, length ≤ 31 chars. */
  readonly name: string;
  /** 0-based position in `XlsxWorkbook.sheets`. */
  readonly index: number;
  /** Hidden state: `"visible" | "hidden" | "veryHidden"`. */
  readonly state: "visible" | "hidden" | "veryHidden";
  /** Sheet kind. Phase 4 records it; commands respect it in Phase 5. */
  readonly kind: "worksheet" | "non-worksheet";
  /**
   * Container path for this sheet's XML
   * (e.g. `xl/worksheets/sheet1.xml`). Resolved via the workbook rels.
   */
  readonly partPath: string;
  /**
   * Container path for this sheet's rels part, if it exists. Used for
   * comment/hyperlink/drawing references.
   */
  readonly relsPartPath?: string;
  /**
   * Typed sparse cell store. Key = `${row}:${col}` with both indices
   * 0-based. Populated by the parser from the SheetJS dense workbook
   * (Phase 5). Phase 5 commands mutate this map; the serializer
   * re-emits the sheet XML when the sheet is in `dirty.sheets`.
   */
  readonly cells: ReadonlyMap<string, Cell>;
  /** Merged regions (rectangular). 0-based inclusive bounds. */
  readonly merges: ReadonlyArray<MergedCell>;
  /**
   * Classic notes anchored to single cells (Phase 7j). Threaded
   * comments (Excel "modern" comments living in
   * `xl/threadedComments/`) remain opaque in P0.
   */
  readonly comments: ReadonlyArray<Comment>;
  /** Path of `xl/comments{N}.xml` if any comments exist. */
  readonly commentsPartPath?: string;
  /** Authors list referenced by classic comments (preserves order). */
  readonly commentAuthors: ReadonlyArray<string>;
  /**
   * Per-column width override in CSS pixels. Key = 0-based column
   * index, value = width in px. Columns not present render at the
   * Grid's default `COL_WIDTH`. Populated by `xlsx:set-column-width`
   * (P11g); P0 round-trip leaves the OOXML `<cols>` band opaque.
   */
  readonly columnWidths: ReadonlyMap<number, number>;
  /**
   * Per-row height override in CSS pixels. Key = 0-based row index,
   * value = height in px. Rows not present render at the Grid's
   * default `ROW_HEIGHT`. Populated by `xlsx:set-row-height` (P11g).
   */
  readonly rowHeights: ReadonlyMap<number, number>;
  /**
   * AutoFilter applied to the sheet, if any. Excel allows at most one
   * `<autoFilter>` per worksheet. The header row is `range.r1`; the
   * remaining rows in the range (`r1+1..r2`) are evaluated against
   * `columns` to populate {@link Sheet.hiddenRows}.
   */
  readonly autoFilter?: AutoFilter;
  /**
   * 0-based row indices currently hidden by the active AutoFilter (or
   * any other "hide row" affordance). The grid renders these rows at
   * height 0 and the serializer emits `hidden="1"` on the matching
   * `<row>` element. Mirrors Excel's filter-driven row hiding —
   * Excel never auto-recomputes hiddenness on cell edits, so we don't
   * either; only filter mutations rebuild this set.
   */
  readonly hiddenRows: ReadonlySet<number>;
  /**
   * Free-floating images anchored to this sheet (raster only in v1).
   * Render order = array order (= z-order: later items overlay
   * earlier ones), matching Excel's drawing list semantics.
   */
  readonly images: ReadonlyArray<SheetImage>;
  /**
   * Container path for this sheet's drawing part
   * (`xl/drawings/drawingN.xml`). Resolved via the sheet's rels.
   * Undefined when the sheet has no images.
   */
  readonly drawingPartPath?: string;
}

/**
 * Excel `<autoFilter>` element. Anchored to a rectangular range whose
 * first row is the header. Per OOXML spec there's at most one per
 * worksheet.
 */
export interface AutoFilter {
  readonly range: AutoFilterRange;
  /**
   * Per-column criteria, keyed by `colId` (0-based offset from
   * `range.c1`). Columns absent from this map carry no filter and
   * never contribute to the row-hide AND.
   */
  readonly columns: ReadonlyMap<number, FilterColumn>;
}

export interface AutoFilterRange {
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
}

/**
 * Discriminated union of every Excel filter criterion shape we
 * support. Mirrors the OOXML `<filterColumn>` children:
 *   - `<filters>`        → `kind: "values"`
 *   - `<customFilters>`  → `kind: "custom"`
 *   - `<top10>`          → `kind: "top10"`
 *   - `<dynamicFilter>`  → `kind: "dynamic"`
 *   - `<colorFilter>`    → `kind: "color"`
 */
export type FilterColumn =
  | FilterColumnValues
  | FilterColumnCustom
  | FilterColumnTop10
  | FilterColumnDynamic
  | FilterColumnColor;

export interface FilterColumnValues {
  readonly kind: "values";
  /**
   * Allowed displayed values. Strings come from the same
   * `formatCellValue` pipeline the grid uses, so "$1,234" matches a
   * cell rendered as currency even though the underlying number is
   * `1234`.
   */
  readonly values: ReadonlySet<string>;
  /** When true, blank cells are also kept (Excel's "(Blanks)" entry). */
  readonly blank: boolean;
}

export interface FilterColumnCustom {
  readonly kind: "custom";
  readonly op1: CustomFilterOp;
  readonly op2?: CustomFilterOp;
  /** Combinator for `op2`. Defaults to "and". */
  readonly combine: "and" | "or";
}

export interface CustomFilterOp {
  readonly operator:
    | "equal"
    | "notEqual"
    | "greaterThan"
    | "greaterThanOrEqual"
    | "lessThan"
    | "lessThanOrEqual";
  /** Comparison value. May contain `*` / `?` wildcards (text only). */
  readonly val: string;
}

export interface FilterColumnTop10 {
  readonly kind: "top10";
  /** True for "Top N" (largest); false for "Bottom N" (smallest). */
  readonly top: boolean;
  /** True for "% of records" mode; false for "Items". */
  readonly percent: boolean;
  /** N — usually 10. Excel exposes 1..500. */
  readonly n: number;
  /** Cached threshold value Excel writes; recomputed by the evaluator. */
  readonly filterVal: number;
}

export interface FilterColumnDynamic {
  readonly kind: "dynamic";
  readonly type: DynamicFilterType;
}

export interface FilterColumnColor {
  readonly kind: "color";
  /** ARGB hex (8 chars, e.g. "FFFFEB9C"). */
  readonly argb: string;
  /** True → fill colour; false → font colour. */
  readonly isCellColor: boolean;
}

/**
 * Excel `<dynamicFilter type="…">` enumeration. Date-relative criteria
 * are evaluated against "today" at filter-apply time.
 */
export type DynamicFilterType =
  | "today"
  | "yesterday"
  | "tomorrow"
  | "thisWeek"
  | "lastWeek"
  | "nextWeek"
  | "thisMonth"
  | "lastMonth"
  | "nextMonth"
  | "thisQuarter"
  | "lastQuarter"
  | "nextQuarter"
  | "thisYear"
  | "lastYear"
  | "nextYear"
  | "yearToDate"
  | "M1"
  | "M2"
  | "M3"
  | "M4"
  | "M5"
  | "M6"
  | "M7"
  | "M8"
  | "M9"
  | "M10"
  | "M11"
  | "M12"
  | "Q1"
  | "Q2"
  | "Q3"
  | "Q4";

export interface Comment {
  readonly id: string;
  /** A1 single-cell ref. */
  readonly ref: string;
  readonly author: string;
  readonly text: string;
  /** Threaded-comment parent id; replies link back to a top-level comment. */
  readonly parentId?: string;
  /**
   * Synthetic resolved flag used by the shared comments UI. Round-trips
   * via the `done` attribute on the corresponding `<threadedComment>`
   * entry so it survives a serialise → re-parse cycle in Excel.
   */
  readonly resolved?: boolean;
  /** ISO-8601 creation timestamp; shown in the comments sidebar. */
  readonly createdAt?: string;
}

export interface Cell {
  /** 0-based row. */
  readonly row: number;
  /** 0-based column. */
  readonly col: number;
  readonly value: CellValue;
  /**
   * Original formula, when present. Phase 5 only authors literal
   * values; setting a formula is `xlsx:set-cell-formula`, which lands
   * in Phase 7 alongside the formula engine. Round-trips verbatim.
   */
  readonly formula?: Formula;
  /**
   * Index into `XlsxWorkbook.styles.cellXfs`. `undefined` = the
   * implicit default xf (id 0). `xlsx:set-cell-format` is the only
   * command that mutates this; other handlers preserve it on the
   * cell and on the SheetJS round-trip via `sheet-sync`.
   */
  readonly styleId?: number;
}

export type CellValue = number | string | boolean | null | CellErrorValue;

export interface CellErrorValue {
  readonly kind: "error";
  /** Excel error sentinel (`#REF!`, `#NAME?`, etc.). */
  readonly code: CellErrorCode;
}

export type CellErrorCode =
  | "#REF!"
  | "#VALUE!"
  | "#DIV/0!"
  | "#NAME?"
  | "#N/A"
  | "#NULL!"
  | "#NUM!"
  | "#GETTING_DATA"
  | "#SPILL!";

export interface Formula {
  /** Original formula text WITHOUT the leading `=`. */
  readonly text: string;
}

export interface MergedCell {
  /** Top row, 0-based inclusive. */
  readonly r1: number;
  /** Left column, 0-based inclusive. */
  readonly c1: number;
  /** Bottom row, 0-based inclusive. */
  readonly r2: number;
  /** Right column, 0-based inclusive. */
  readonly c2: number;
}

export interface OpaquePart {
  /** Full zip path, e.g. `xl/charts/chart1.xml`. */
  readonly path: string;
  readonly bytes: Uint8Array;
  /** Content type from `[Content_Types].xml` when known. */
  readonly contentType?: string;
  /** SHA-256 hex digest of `bytes`. */
  readonly hash: string;
}

/**
 * Empty dirty-flags constant, returned by `parseXlsx`. Helper to keep
 * the parser focussed on parsing.
 */
export function emptyDirty(): XlsxDirtyFlags {
  return {
    workbook: false,
    sharedStrings: false,
    styles: false,
    contentTypes: false,
    rels: false,
    sheets: new Set<string>(),
    comments: new Set<string>(),
    threadedComments: new Set<string>(),
    sheetRels: new Set<string>(),
    removedSheetParts: new Set<string>(),
    drawings: new Set<string>(),
    media: new Set<string>(),
    removedMediaParts: new Set<string>(),
  };
}
