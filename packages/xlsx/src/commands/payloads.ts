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
