import type { ImageContentType } from "../model/drawings.js";
import type {
  CellErrorCode,
  CellValue,
  CustomFilterOp,
  DynamicFilterType,
  FilterColumn,
} from "../model/types.js";

/**
 * Wire payloads for every `xlsx:*` command. Mirrors
 * `spec/xlsx/agent-commands.md`. JSON-serializable. Values are typed
 * against the model's `CellValue` union; agents pass `{ kind: "error",
 * code: "#REF!" }` for error sentinels.
 */

export type { CellErrorCode, CellValue, CustomFilterOp, DynamicFilterType, FilterColumn };

/** `xlsx:set-cell-value` */
export interface SetCellValuePayload {
  readonly sheet: string;
  /** A1 single-cell ref, e.g. `"B2"`. */
  readonly ref: string;
  readonly value: CellValue;
}

/** `xlsx:set-cell-formula` */
export interface SetCellFormulaPayload {
  readonly sheet: string;
  /** A1 single-cell ref, e.g. `"C2"`. */
  readonly ref: string;
  /** Formula text, with or without leading `=`. Empty body clears the cell. */
  readonly formula: string;
}

/** `xlsx:set-range-values` */
export interface SetRangeValuesPayload {
  readonly sheet: string;
  /** A1 range, e.g. `"A1:C3"`. */
  readonly range: string;
  /** Row-major 2-D matrix; dimensions MUST equal the range. */
  readonly values: ReadonlyArray<ReadonlyArray<CellValue>>;
}

/** `xlsx:merge-cells` */
export interface MergeCellsPayload {
  readonly sheet: string;
  /** A1 range covering ≥ 2 cells, e.g. `"A1:C1"`. */
  readonly range: string;
}

/** `xlsx:unmerge-cells` */
export interface UnmergeCellsPayload {
  readonly sheet: string;
  /** Must exactly match an existing merge range. */
  readonly range: string;
}

/** `xlsx:add-sheet` */
export interface AddSheetPayload {
  readonly name: string;
  /** 0-based insert position; defaults to append (= `sheets.length`). */
  readonly at?: number;
}

/** `xlsx:rename-sheet` */
export interface RenameSheetPayload {
  /** Current sheet name (case-sensitive lookup). */
  readonly name: string;
  /** New sheet name. Validated against Excel naming rules. */
  readonly newName: string;
}

/**
 * `xlsx:move-sheet` — reorder a worksheet within the workbook.
 *
 * The serializer rewrites the workbook `<sheets>` block from the new
 * `workbook.sheets` order, so the OOXML round-trip is preserved.
 * Cross-sheet formula refs are unaffected (Excel anchors them by
 * sheet name, not position).
 */
export interface MoveSheetPayload {
  /** Current sheet name (case-sensitive lookup). */
  readonly name: string;
  /**
   * Target 0-based position in the final ordering. Clamped to
   * `[0, sheets.length-1]`. The currently-active sheet stays
   * selected because the editor keys on `name`, not index.
   */
  readonly to: number;
}

/**
 * `xlsx:set-sheet-state` — flip a sheet between
 * `visible` / `hidden` / `veryHidden`. Excel uses this for the
 * "Hide" / "Unhide…" affordance on the sheet tab context menu.
 *
 * The command refuses to leave the workbook with zero visible
 * sheets (Excel parity).
 */
export interface SetSheetStatePayload {
  readonly name: string;
  readonly state: "visible" | "hidden" | "veryHidden";
}

/**
 * `xlsx:add-conditional-format` — add a typed conditional
 * formatting rule to a sheet.
 *
 * Round-trip caveat: typed CF rules are evaluated at render time
 * but not yet emitted to OOXML on serialize. Existing CF blocks
 * from the parsed workbook are preserved verbatim.
 */
export interface AddConditionalFormatPayload {
  readonly sheet: string;
  readonly rule: import("../model/types.js").ConditionalFormat;
}

/**
 * `xlsx:remove-conditional-format` — drop a single rule by id.
 */
export interface RemoveConditionalFormatPayload {
  readonly sheet: string;
  readonly id: string;
}

/**
 * `xlsx:clear-conditional-formats` — drop every typed CF rule on a
 * sheet. Opaque preserved blocks are NOT cleared by this command.
 */
export interface ClearConditionalFormatsPayload {
  readonly sheet: string;
}

/**
 * `xlsx:add-data-validation` — append a typed data-validation rule
 * (`list` only in C11; richer kinds preserve via opaque round-trip
 * until we extend the model).
 */
export interface AddDataValidationPayload {
  readonly sheet: string;
  readonly rule: import("../model/types.js").DataValidation;
}

/**
 * `xlsx:remove-data-validation` — drop a single typed rule by id.
 * Opaque (non-list) rules captured at parse time aren't addressable
 * from this command; clear them via the dialog's "Remove all rules"
 * action which calls `clear-data-validations`.
 */
export interface RemoveDataValidationPayload {
  readonly sheet: string;
  readonly id: string;
}

/** `xlsx:clear-data-validations` — wipes both typed and opaque rules. */
export interface ClearDataValidationsPayload {
  readonly sheet: string;
}

/* ── Defined names (named ranges) — `xlsx:add-defined-name` ───────────────
 * C12. Workbook- or sheet-scoped names that resolve to a range / cell /
 * formula expression. Round-trips through `xl/workbook.xml`'s
 * `<definedNames>` block.
 */

export interface AddDefinedNamePayload {
  /** Display name, e.g. `"Revenue"`. Excel rules: starts with letter or `_`,
   * no spaces, ≤ 255 chars, not a cell ref (`A1`/`R1C1`), not a reserved word. */
  readonly name: string;
  /** OOXML refersTo expression without leading `=`, e.g. `"Sheet1!$A$1:$B$5"`. */
  readonly refersTo: string;
  /** Sheet-name scope. Omitted = workbook-scoped. */
  readonly scope?: string;
  readonly comment?: string;
}

export interface UpdateDefinedNamePayload {
  /** Lookup key — name + scope identify the entry uniquely. */
  readonly name: string;
  readonly scope?: string;
  /** New name (rename). When omitted, name stays the same. */
  readonly nextName?: string;
  /** New refersTo expression. */
  readonly refersTo?: string;
  readonly comment?: string;
}

export interface RemoveDefinedNamePayload {
  readonly name: string;
  readonly scope?: string;
}

/* ── `xlsx:set-cell-format` (§4) ──────────────────────────────────────────
 * Patch-style format payload — undefined fields are left unchanged.
 * Per spec/xlsx/agent-commands.md §4, `format` carries the friendly
 * agent-facing names (`bold`, `RRGGBB` colours, etc.); the handler
 * translates to the OOXML style table shape internally.
 */

export interface CellFormatBorderSide {
  readonly style?: "thin" | "medium" | "thick" | "double" | "dashed" | "dotted" | "none";
  /** RRGGBB hex without `#`. */
  readonly color?: string;
}

export interface CellFormatPatch {
  readonly font?: {
    /**
     * Font family name (e.g. "Calibri"). Maps to the OOXML `<name>`
     * element on `<font>`. The legacy alias `family` is still
     * accepted for backwards compatibility.
     */
    readonly fontFamily?: string;
    /** @deprecated Use `fontFamily`. Kept so existing call sites keep working. */
    readonly family?: string;
    /** Font size in points. */
    readonly size?: number;
    readonly bold?: boolean;
    readonly italic?: boolean;
    readonly underline?: boolean;
    readonly strike?: boolean;
    /** RRGGBB hex without `#`. */
    readonly color?: string;
  };
  readonly fill?: {
    /** RRGGBB hex without `#`. */
    readonly color?: string;
    readonly pattern?: "solid" | "none";
  };
  readonly border?: {
    readonly top?: CellFormatBorderSide;
    readonly right?: CellFormatBorderSide;
    readonly bottom?: CellFormatBorderSide;
    readonly left?: CellFormatBorderSide;
  };
  readonly alignment?: {
    readonly horizontal?: "left" | "center" | "right" | "fill" | "justify";
    readonly vertical?: "top" | "middle" | "bottom";
    readonly wrapText?: boolean;
    readonly indent?: number;
  };
  /** Built-in numFmtId as a string, or a custom format string. */
  readonly numberFormat?: string;
  /**
   * Cell protection flags. Honored only when the worksheet itself is
   * protected — Excel's "lock" attribute is a per-cell hint, but the
   * protection only kicks in once the sheet is locked. We model both
   * fields explicitly so a Format Cells round-trip preserves whatever
   * the source workbook had.
   */
  readonly protection?: {
    readonly locked?: boolean;
    readonly hidden?: boolean;
  };
}

export interface SetCellFormatPayload {
  readonly sheet: string;
  /** A1 single cell or A1 range, e.g. `"B2"` or `"A1:E1"`. */
  readonly range: string;
  readonly format: CellFormatPatch;
}

/* ── Structural reshape (§§5–8) ───────────────────────────────────────────
 * `at` is 1-based to match A1 row / column indexing; `count` ≥ 1.
 */

/** `xlsx:insert-row` */
export interface InsertRowPayload {
  readonly sheet: string;
  /** 1-based row index; insertion is BEFORE this row. */
  readonly at: number;
  /** Number of blank rows to insert. Must satisfy `at + count - 1 ≤ 1048576`. */
  readonly count: number;
}

/** `xlsx:insert-column` */
export interface InsertColumnPayload {
  readonly sheet: string;
  /** 1-based column index (A=1); insertion is to the LEFT of this column. */
  readonly at: number;
  /** Number of blank columns to insert. Must satisfy `at + count - 1 ≤ 16384`. */
  readonly count: number;
}

/** `xlsx:delete-row` */
export interface DeleteRowPayload {
  readonly sheet: string;
  /** 1-based row index of the first row to drop. */
  readonly at: number;
  /** Number of rows to drop. Removes rows `at..at+count-1`. */
  readonly count: number;
}

/** `xlsx:delete-column` */
export interface DeleteColumnPayload {
  readonly sheet: string;
  /** 1-based column index (A=1) of the first column to drop. */
  readonly at: number;
  /** Number of columns to drop. Removes columns `at..at+count-1`. */
  readonly count: number;
}

/* ── `xlsx:set-column-width` / `xlsx:set-row-height` (P11g) ──────────────
 * Per-column / per-row size overrides in CSS pixels. The handlers
 * mutate `sheet.columnWidths` / `sheet.rowHeights`; OOXML round-trip
 * leaves `<cols>` / `<row ht=…>` opaque in P0 — these sizes are a
 * runtime UI affordance for the web grid and the diff trail, not yet
 * a formal serializer-side feature.
 */

/** `xlsx:set-column-width` */
export interface SetColumnWidthPayload {
  readonly sheet: string;
  /** 1-based column index (A=1). */
  readonly column: number;
  /** Width in CSS pixels. Pass `null` to reset to the default. */
  readonly width: number | null;
}

/** `xlsx:set-row-height` */
export interface SetRowHeightPayload {
  readonly sheet: string;
  /** 1-based row index. */
  readonly row: number;
  /** Height in CSS pixels. Pass `null` to reset to the default. */
  readonly height: number | null;
}

/** `xlsx:delete-sheet` */
export interface DeleteSheetPayload {
  /** Sheet name to drop (case-sensitive lookup). */
  readonly name: string;
}

/** `xlsx:add-comment` */
export interface AddCommentPayload {
  readonly sheet: string;
  /** A1 single-cell ref, e.g. `"B7"`. Range refs are rejected. */
  readonly ref: string;
  /** Comment body. Plain text in P0 (rich-text formatting deferred). */
  readonly text: string;
  readonly author: string;
}

/** `xlsx:reply-comment` (§13) — appends a reply to an existing thread. */
export interface ReplyCommentPayload {
  readonly sheet: string;
  /** Top-level comment id this reply attaches to. */
  readonly parentId: string;
  readonly author: string;
  readonly text: string;
}

/** `xlsx:resolve-comment` (§13) — toggles a thread's resolved flag. */
export interface ResolveCommentPayload {
  readonly sheet: string;
  readonly commentId: string;
  readonly resolved: boolean;
}

/** `xlsx:delete-comment` (§13) — removes a comment (and any replies if top-level). */
export interface DeleteCommentPayload {
  readonly sheet: string;
  readonly commentId: string;
}

/** `xlsx:edit-comment` (§13) — rewrites the plain-text body of a comment. */
export interface EditCommentPayload {
  readonly sheet: string;
  readonly commentId: string;
  readonly text: string;
}

/* ── P13: Clipboard / Fill / Text-to-Columns ─────────────────────────────── */

/**
 * `xlsx:paste-range` (§14)
 *
 * Atomically writes a {@link XlsxClipboardSnapshot} at the target
 * top-left cell. Modes:
 *   - `"all"`     — values + formulas + styleIds + merges (default)
 *   - `"values"`  — values + formulas only (style preserved)
 *   - `"formats"` — styleIds only (value preserved)
 *
 * `transpose: true` flips rows and columns before applying.
 */
export interface PasteRangePayload {
  readonly sheet: string;
  /** A1 single-cell ref pointing at the destination top-left. */
  readonly target: string;
  readonly source: import("../clipboard/snapshot.js").XlsxClipboardSnapshot;
  /**
   * Paste Special variants:
   *   - "all"      : values + formulas + formats (default; current Cmd+V)
   *   - "values"   : computed values only — formulas collapse to their
   *                  cached value, no styles, no merges
   *   - "formulas" : formulas (relative-shifted) + cached values for
   *                  literals, no styles, no merges
   *   - "formats"  : per-cell style id only, never overwrites values
   */
  readonly mode?: "all" | "values" | "formulas" | "formats";
  readonly transpose?: boolean;
}

/**
 * `xlsx:fill-range` (§15) — Excel's drag-the-corner fill handle.
 *
 * `source` is the original selection (the cells the user grabbed);
 * `target` extends it by one direction. The handler diffs the two
 * rectangles, picks a series detector against the source data, and
 * writes the extrapolated cells.
 */
export interface FillRangePayload {
  readonly sheet: string;
  readonly source: string;
  readonly target: string;
  readonly direction: "down" | "right" | "up" | "left";
}

/**
 * `xlsx:text-to-columns` (§16) — split each row of `range` on a
 * delimiter and write the columns out in row-major order starting
 * at `destination` (defaults to the same top-left as `range`).
 */
export interface TextToColumnsPayload {
  readonly sheet: string;
  readonly range: string;
  readonly delimiter: string;
  readonly treatConsecutiveAsOne?: boolean;
  readonly destination?: string;
}

/* ── AutoFilter (§17) ────────────────────────────────────────────────────
 * Per-column header dropdowns with sort, value checklist, and
 * text / number / date / colour condition filters. Round-trips through
 * the worksheet's native `<autoFilter>` element so the saved file
 * opens in Excel with the same filter applied.
 */

/**
 * `xlsx:set-auto-filter` — toggle the AutoFilter band on a sheet.
 *
 * Pass `range: null` to remove the filter entirely (and unhide every
 * filter-driven hidden row). Setting a fresh range clears any
 * pre-existing per-column criteria.
 */
export interface SetAutoFilterPayload {
  readonly sheet: string;
  /** A1 range covering header + body, e.g. `"A1:E100"`. `null` removes. */
  readonly range: string | null;
}

/**
 * `xlsx:set-filter-column` — set / replace the criterion on one column
 * of the active AutoFilter. The handler recomputes hidden rows.
 */
export interface SetFilterColumnPayload {
  readonly sheet: string;
  /** 0-based offset from `autoFilter.range.c1`. */
  readonly colId: number;
  readonly criterion: FilterColumn;
}

/**
 * `xlsx:clear-filter-column` — drop the criterion on one column;
 * recomputes hidden rows.
 */
export interface ClearFilterColumnPayload {
  readonly sheet: string;
  readonly colId: number;
}

/**
 * `xlsx:sort-range` — sort the rows inside `range` (excluding the
 * header) by `sortBy.colId`. Used by both standalone Sort and the
 * dropdown's Sort A→Z / Z→A. Mutates cells in place via the existing
 * `set-range-values` plumbing; no autoFilter mutation.
 */
export interface SortRangePayload {
  readonly sheet: string;
  /** A1 range; the first row is treated as the header and never moved. */
  readonly range: string;
  readonly sortBy: {
    /** 0-based offset from the range's first column. */
    readonly colId: number;
    readonly order: "asc" | "desc";
  };
}

/* ── Images (raster, free-floating overlays) ─────────────────────────────
 * v1 only authors `editAs="oneCell"` — images move with cells but do
 * not size with cells. Anchor + dimensions live in CSS pixels in the
 * model; the serializer converts to EMUs.
 */

/** `xlsx:add-image` */
export interface AddImagePayload {
  readonly sheet: string;
  readonly bytes: Uint8Array;
  readonly contentType: ImageContentType;
  /** Optional friendly name. Defaults to `Picture N`. */
  readonly name?: string;
  readonly altText?: string;
  /** Top-left anchor: 0-based row + column the image is pinned to. */
  readonly fromRow: number;
  readonly fromCol: number;
  /** Pixel offset INTO the from-cell. Defaults to 0,0. */
  readonly fromOffsetXPx?: number;
  readonly fromOffsetYPx?: number;
  /** Rendered size in CSS pixels. Defaults to the natural image size. */
  readonly widthPx: number;
  readonly heightPx: number;
}

/** `xlsx:move-image` — repin an image to a new from-cell + offset. */
export interface MoveImagePayload {
  readonly sheet: string;
  readonly imageId: string;
  readonly fromRow: number;
  readonly fromCol: number;
  readonly fromOffsetXPx: number;
  readonly fromOffsetYPx: number;
}

/** `xlsx:resize-image` — change the rendered width/height (in CSS pixels). */
export interface ResizeImagePayload {
  readonly sheet: string;
  readonly imageId: string;
  readonly widthPx: number;
  readonly heightPx: number;
}

/** `xlsx:remove-image` — drop an image from a sheet (GCs orphan media). */
export interface RemoveImagePayload {
  readonly sheet: string;
  readonly imageId: string;
}

/* ── Freeze panes (C3) ─────────────────────────────────────────────────── */

/**
 * `xlsx:freeze-panes` — freeze a number of rows from the top and/or
 * columns from the left. Setting both to `0` is equivalent to
 * `xlsx:unfreeze-panes`.
 *
 * Bounds:
 *   - `rows` ∈ [0, 1048576]
 *   - `cols` ∈ [0, 16384]
 *
 * Excel parity: any non-zero `rows`/`cols` writes a
 * `<pane state="frozen"/>` block to the worksheet's `<sheetView>`,
 * which Excel restores on open. Round-trips through our serializer.
 */
export interface FreezePanesPayload {
  readonly sheet: string;
  readonly rows: number;
  readonly cols: number;
}

/** `xlsx:unfreeze-panes` — clear any frozen-pane configuration. */
export interface UnfreezePanesPayload {
  readonly sheet: string;
}

/* ── Excel Tables / ListObjects (C14) ──────────────────────────────────── */

/**
 * `xlsx:add-table` — promote a range to an Excel Table.
 *
 * Mirrors Excel's Ctrl+T affordance:
 *   - `range` covers headers + body. The first row is treated as the
 *     header row when `hasHeaders !== false`.
 *   - When `hasHeaders` is `false`, the table gets synthesised
 *     `Column1`, `Column2`, … names just like Excel does.
 *   - The table's name is derived from `name` if supplied, otherwise
 *     the next available `TableN` (workbook-scoped, since OOXML
 *     enforces global table-name uniqueness).
 *
 * The handler also installs an AutoFilter at the table range so
 * filter buttons appear immediately, matching Excel's default. C14
 * does not (yet) auto-extend the table on adjacent edits — that
 * lives in a follow-up.
 */
export interface AddTablePayload {
  readonly sheet: string;
  /** A1 range, e.g. `"A1:E25"`. Must contain at least one body row. */
  readonly range: string;
  /** Defaults to `true`. */
  readonly hasHeaders?: boolean;
  /** Optional explicit table name. Must be unique workbook-wide. */
  readonly name?: string;
}

/** `xlsx:remove-table` — remove an Excel Table from a sheet (cells stay). */
export interface RemoveTablePayload {
  readonly sheet: string;
  readonly tableId: string;
}

/* ── Charts (C15) ──────────────────────────────────────────────────────── */

/**
 * `xlsx:add-chart` — drop a free-floating chart on a worksheet.
 *
 * MVP scope: column / bar / line / pie. The chart auto-anchors near
 * the right edge of `dataRange` so it doesn't cover the source data,
 * unless an explicit `anchor` is provided.
 */
export interface AddChartPayload {
  readonly sheet: string;
  readonly kind: import("../model/types.js").ChartKind;
  /** A1 range of the chart's data, e.g. `"A1:B7"`. */
  readonly dataRange: string;
  /** Defaults to `true` — first row becomes the series legend. */
  readonly hasHeaderRow?: boolean;
  /** Defaults to `true` — first column becomes the category labels. */
  readonly hasCategoryColumn?: boolean;
  readonly title?: string;
  /** Optional explicit anchor; auto-derived from `dataRange` when omitted. */
  readonly anchor?: import("../model/drawings.js").ImageAnchor;
  /** Series color theme. Defaults to `"default"` when omitted. */
  readonly palette?: import("../model/types.js").ChartPalette;
  /** Default `true`. */
  readonly showLegend?: boolean;
  /** Default `false`. */
  readonly showDataLabels?: boolean;
  /** Default `true`. */
  readonly showGridlines?: boolean;
  readonly xAxisTitle?: string;
  readonly yAxisTitle?: string;
}

/** `xlsx:remove-chart` — drop a chart from a sheet. */
export interface RemoveChartPayload {
  readonly sheet: string;
  readonly chartId: string;
}

/** `xlsx:move-chart` — repin a chart's anchor without resizing. */
export interface MoveChartPayload {
  readonly sheet: string;
  readonly chartId: string;
  readonly fromRow: number;
  readonly fromCol: number;
  readonly fromOffsetXPx: number;
  readonly fromOffsetYPx: number;
}

/** `xlsx:resize-chart` — change a chart's rendered size in CSS pixels. */
export interface ResizeChartPayload {
  readonly sheet: string;
  readonly chartId: string;
  readonly widthPx: number;
  readonly heightPx: number;
}

/**
 * `xlsx:update-chart` — patch a chart's typed properties in place.
 *
 * Every field is optional; only provided fields are written. `title`
 * uses an explicit `null` to mean "remove the title" (vs. `undefined`
 * = "leave alone") because the model uses an optional string. At least
 * one mutable property must be supplied or the handler rejects the
 * command — silent no-ops are the agent equivalent of a typo.
 *
 * Position / size are deliberately routed through the existing
 * `xlsx:move-chart` and `xlsx:resize-chart` commands so undo/redo
 * stays a single step per logical gesture.
 */
export interface UpdateChartPayload {
  readonly sheet: string;
  readonly chartId: string;
  readonly kind?: import("../model/types.js").ChartKind;
  /** A1 range; same validation as `xlsx:add-chart` (≥ 2 cells). */
  readonly dataRange?: string;
  /** `null` clears the title; `undefined` leaves it unchanged. */
  readonly title?: string | null;
  readonly hasHeaderRow?: boolean;
  readonly hasCategoryColumn?: boolean;
  /**
   * Style fields. As with `title`, `null` resets the field to its
   * default (renderer-side) and `undefined` leaves it untouched —
   * `palette: null` reverts to `"default"`, the boolean toggles
   * revert to their renderer defaults, and the axis titles clear.
   */
  readonly palette?: import("../model/types.js").ChartPalette | null;
  readonly showLegend?: boolean | null;
  readonly showDataLabels?: boolean | null;
  readonly showGridlines?: boolean | null;
  readonly xAxisTitle?: string | null;
  readonly yAxisTitle?: string | null;
}
