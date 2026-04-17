import type { DocumentSnapshot, NodeId, ooxml } from "@officeai/core";

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
   * Phase 4 escape hatch: the SheetJS `WorkBook` produced by
   * `XLSX.read(..., { dense: true, ... })` carrying the cell layer.
   * Phase 5 introduces a typed `cells: Map<string, Cell>` on each
   * `Sheet` and this field becomes derived. Consumers should NOT take
   * a hard dependency on its shape outside `parser/`, `serializer/`,
   * and the Phase 5 model upgrade path.
   */
  readonly sheetjs: import("xlsx").WorkBook;
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
  };
}
