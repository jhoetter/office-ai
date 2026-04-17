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
