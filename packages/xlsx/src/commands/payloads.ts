import type { CellErrorCode, CellValue } from "../model/types.js";

/**
 * Wire payloads for every `xlsx:*` command. Mirrors
 * `spec/xlsx/agent-commands.md`. JSON-serializable. Values are typed
 * against the model's `CellValue` union; agents pass `{ kind: "error",
 * code: "#REF!" }` for error sentinels.
 */

export type { CellErrorCode, CellValue };

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
    readonly family?: string;
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
  readonly mode?: "all" | "values" | "formats";
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
