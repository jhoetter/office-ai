# XLSX — In-Memory Model

> The format-specific projection of the shared `DocumentSnapshot<TRoot>`
> contract for `@officeai/xlsx`. Every typed shape an agent, command
> handler, renderer, or formula engine touches is declared here.
>
> Companion files: [`ooxml-mapping.md`](ooxml-mapping.md) (which OOXML
> element/attribute populates each field), [`feature-scope.md`](feature-scope.md)
> (which fields the agent may mutate in P0), and the shared
> [`spec/shared/document-model.md`](../shared/document-model.md)
> (the `DocumentSnapshot` / `NodeId` / `DocumentDiff` envelope this
> file extends).
>
> This doc is self-contained: a reader who has never opened the Univer
> or SheetJS source must be able to implement the parser, mutation
> layer, and serializer from this file plus `ooxml-mapping.md`.

## 1. Roots

```typescript
import type { DocumentSnapshot, NodeId, ooxml } from "@officeai/core";

/**
 * The XLSX projection of `DocumentSnapshot`. Mirrors `DocxSnapshot` in
 * shape: a `format` discriminator, a `root` carrying the typed model, a
 * `revision` counter (bumped by every applied mutation), `partHashes`
 * keyed by full zip path (drives byte-preservation in the serializer),
 * an internal `container` holding original bytes for untouched parts,
 * and per-part `dirty` flags consulted by the serializer.
 */
export interface XlsxSnapshot extends DocumentSnapshot<XlsxWorkbook> {
  readonly format: "xlsx";
  /** Internal: the OOXML container. Not serialized; not part of equality. */
  readonly container: ooxml.OoxmlContainer;
  /** Per-part dirty flags consulted by the serializer (see §10). */
  readonly dirty: XlsxDirtyFlags;
}

export interface XlsxDirtyFlags {
  /** `xl/workbook.xml` (sheet list, defined names, workbook properties). */
  workbook: boolean;
  /** `xl/sharedStrings.xml`. */
  sharedStrings: boolean;
  /** `xl/styles.xml`. */
  styles: boolean;
  /** `[Content_Types].xml`. */
  contentTypes: boolean;
  /** `_rels/.rels` and `xl/_rels/workbook.xml.rels`. */
  rels: boolean;
  /**
   * Per-sheet dirtiness. Key = sheet part path (`xl/worksheets/sheet{N}.xml`).
   * A sheet not in the set re-emits from its cached bytes — the
   * byte-preservation invariant.
   */
  sheets: ReadonlySet<string>;
  /**
   * Per-comments-part dirtiness. Key = part path
   * (`xl/comments{N}.xml`). Modern threaded-comments part
   * (`xl/threadedComments/threadedComment{N}.xml`) is tracked
   * separately under `threadedComments`.
   */
  comments: ReadonlySet<string>;
  threadedComments: ReadonlySet<string>;
  /** Per-sheet rels parts (`xl/worksheets/_rels/sheet{N}.xml.rels`). */
  sheetRels: ReadonlySet<string>;
}
```

## 2. Workbook

```typescript
export interface XlsxWorkbook {
  readonly id: NodeId;
  /** Sheets in tab order. Order matters; `index` mirrors the array index. */
  readonly sheets: ReadonlyArray<Sheet>;
  /** Workbook-level shared style table. */
  readonly styleTable: StyleTable;
  /** Shared-strings table. */
  readonly sharedStrings: SharedStringTable;
  /** Workbook-level + sheet-level defined names. */
  readonly definedNames: ReadonlyArray<DefinedName>;
  /**
   * Hash of every OOXML part as it was on parse. Same shape as
   * `DocumentSnapshot.partHashes`. Doubly-stored here only because the
   * serializer reads from the workbook root.
   */
  readonly partHashes: Readonly<Record<string, string>>;
  /**
   * Original namespace declarations + non-content root attributes from
   * `xl/workbook.xml`. Re-emitted verbatim when the workbook part is
   * dirty.
   */
  readonly workbookRootAttrs: Readonly<Record<string, string>>;
  /**
   * `WorkbookPr.date1904` flag. When `true`, all date serials are
   * relative to 1904-01-01 instead of 1900-01-00. We do not normalize:
   * the flag travels with the workbook and is consumed by the
   * date-coercion layer at the boundary.
   */
  readonly date1904: boolean;
  /**
   * Theme color palette resolved from `xl/theme/theme1.xml`. Other theme
   * parts (theme2.xml, theme3.xml, …) live in the container as opaque
   * parts and are not parsed in P0.
   */
  readonly theme?: WorkbookTheme;
  /**
   * Opaque parts that the model does not type. Indexed by full zip path
   * for fast lookup during round-trip and during operations that need
   * to mint new top-level parts (charts, drawings, embeddings, …). Each
   * entry round-trips byte-identical.
   */
  readonly opaqueParts: ReadonlyMap<string, OpaquePart>;
}
```

### 2.1 `WorkbookTheme`

```typescript
/** Subset of theme color scheme actually consumed by the renderer. */
export interface WorkbookTheme {
  /** Indexable theme colors in OOXML's canonical order. Six entries. */
  readonly colors: ReadonlyArray<{
    readonly name: string; // "lt1", "dk1", "accent1", …
    readonly rgb: string; // 6-hex, no leading "#"
  }>;
  /** The original `theme1.xml` bytes in the container so unmodelled
   * elements (font scheme, format scheme, custom color stops) round-trip. */
  readonly partPath: "xl/theme/theme1.xml";
}
```

### 2.2 `OpaquePart`

```typescript
/**
 * An OOXML part the model does not type. Stored verbatim and round-
 * tripped byte-identical. The `bytes` field is a reference into
 * `OoxmlContainer.parts`; mutations create new instances rather than
 * editing in place.
 */
export interface OpaquePart {
  /** Full zip path, e.g. `xl/charts/chart1.xml`. */
  readonly path: string;
  readonly bytes: Uint8Array;
  /** Content type from `[Content_Types].xml`. */
  readonly contentType: string;
  /** SHA-256 hex digest of `bytes`. */
  readonly hash: string;
}
```

## 3. Sheet

```typescript
export interface Sheet {
  readonly id: NodeId;
  /** Stable OOXML sheet id (`<sheet sheetId="…">`). String, not the index. */
  readonly sheetId: string;
  /** Display name as it appears on the tab. UTF-8, length ≤ 31 chars (Excel limit). */
  readonly name: string;
  /** 0-based position in `XlsxWorkbook.sheets`. Mirrored from the array index. */
  readonly index: number;
  /** Tab color (RRGGBB, no leading `#`) or undefined when not set. */
  readonly tabColor?: string;
  /** Hidden state: `"visible" | "hidden" | "veryHidden"`. */
  readonly hidden?: "hidden" | "veryHidden";
  /**
   * Frozen-pane split row: number of rows kept stationary at the top.
   * 0-based count, so `frozenRows: 1` freezes the first row.
   * Undefined when no horizontal freeze is active.
   */
  readonly frozenRows?: number;
  /** Frozen-pane split column. Same convention as `frozenRows`. */
  readonly frozenCols?: number;
  /** Sparse cell store. Key = `${row}:${col}` with row & col both 0-based. */
  readonly cells: ReadonlyMap<string, Cell>;
  /** Merged regions (rectangular). */
  readonly merges: ReadonlyArray<MergedCell>;
  /** Conditional-formatting rules in priority order. */
  readonly conditionalFormats: ReadonlyArray<ConditionalFormatRule>;
  /** Hyperlinks — referenced by `Cell.hyperlinkId`. */
  readonly hyperlinks: ReadonlyArray<Hyperlink>;
  /** Comments / notes — referenced by `Cell.commentId`. */
  readonly comments: ReadonlyArray<Comment>;
  /** Auto-filter range, if any. Conditions preserved opaquely. */
  readonly autoFilter?: AutoFilter;
  /**
   * Used range, computed from the cells map. `r1/c1` is the top-left,
   * `r2/c2` the bottom-right (both inclusive, 0-based). All-zero when
   * the sheet is empty.
   */
  readonly dimensions: { readonly r1: number; readonly c1: number; readonly r2: number; readonly c2: number };
  /**
   * Per-row metadata (height, hidden, outline level, custom format).
   * Sparse — only rows with non-default state appear. Key = 0-based
   * row index.
   */
  readonly rows: ReadonlyMap<number, RowProperties>;
  /**
   * Per-column metadata. Sparse. Key = 0-based column index.
   */
  readonly cols: ReadonlyMap<number, ColumnProperties>;
  /**
   * Original sheet-XML root attributes (namespace declarations, etc.)
   * captured from the `<worksheet>` element. Re-emitted verbatim.
   */
  readonly sheetRootAttrs: Readonly<Record<string, string>>;
  /**
   * Catch-all bag of unmodelled top-level worksheet children (e.g.
   * `<dataValidations>`, `<sheetProtection>`, `<pageMargins>`,
   * `<headerFooter>`, `<extLst>`). Stored as raw XML fragments so the
   * serializer can splice them back into the worksheet at their
   * canonical position. Round-trips byte-identical when the sheet is
   * not dirty.
   */
  readonly opaqueWorksheetChildren: ReadonlyArray<OpaqueXml>;
  /**
   * Sheet kind. P0 only models worksheets. Chartsheets / dialogsheets
   * / macrosheets parse to `kind: "non-worksheet"` with their full
   * bytes preserved opaquely; the agent surface refuses edits to them.
   */
  readonly kind: "worksheet" | "non-worksheet";
}

export interface RowProperties {
  /** Height in points (`ht`). Undefined = workbook default. */
  readonly height?: number;
  readonly hidden?: boolean;
  /** Outline level (`outlineLevel`), 0-7. */
  readonly outlineLevel?: number;
  /** Default style id applied to the entire row when set. */
  readonly styleId?: number;
  /** Whether the row carries a custom (non-default) format. */
  readonly customFormat?: boolean;
}

export interface ColumnProperties {
  /** Width in character units (`width`). Undefined = workbook default. */
  readonly width?: number;
  readonly hidden?: boolean;
  readonly outlineLevel?: number;
  readonly styleId?: number;
  readonly customWidth?: boolean;
}

export interface AutoFilter {
  readonly ref: RangeRef;
  /**
   * Filter conditions, preserved opaquely in P0. The agent may apply
   * or clear an auto-filter (which sets/clears `ref`) but does not
   * author conditions; existing conditions round-trip verbatim.
   */
  readonly opaqueConditions: ReadonlyArray<OpaqueXml>;
}
```

### 3.1 Sparse cell store

The `Sheet.cells` map is keyed by a fixed string format
`${row}:${col}` with both indices 0-based and stringified without
zero-padding.

> NOTE: We use a `Map<string, Cell>` rather than a 2-level object
> (`{ [row]: { [col]: Cell } }`). Justification: agent workloads write
> in scattered patterns (a few cells across many rows) where the
> 2-level object's intermediate row-objects waste memory; the
> `Map<string, Cell>` is also faster for `delete` and for iteration
> via `for…of`. The address-encoding cost (`row * X + col`) was
> rejected because `X` is not bounded — a sheet may be 1,048,576 rows.

Helpers exposed by `packages/xlsx/src/model/cells.ts`:

```typescript
/** Encode a (row, col) into a cells-map key. */
export function cellKey(row: number, col: number): string;
/** Decode a cells-map key into (row, col). */
export function parseCellKey(key: string): { row: number; col: number };
```

## 4. Cell

```typescript
export interface Cell {
  readonly row: number; // 0-based
  readonly col: number; // 0-based
  readonly value: CellValue;
  /**
   * Original formula. When present, `value` holds the cached evaluation
   * result (the value Excel last wrote into `<v>`). When the formula
   * has not been evaluated yet by our engine, `value` is `null` and
   * the renderer shows the cached display string (or `#GETTING_DATA`
   * if missing).
   */
  readonly formula?: Formula;
  /**
   * Index into `XlsxWorkbook.styleTable.cellXfs`. Undefined =
   * default style (xfId 0).
   */
  readonly styleId?: number;
  /** Index into `Sheet.hyperlinks`. */
  readonly hyperlinkId?: number;
  /** Index into `Sheet.comments`. */
  readonly commentId?: number;
  /**
   * True when the cell carries Excel-365 dynamic-array spill metadata
   * (the `cm="1"` attribute pointing at an `XLDAPR` block in
   * `xl/metadata.xml`). Treated as a flag: the metadata block itself
   * is round-tripped opaquely.
   */
  readonly dynamicArray?: boolean;
}
```

### 4.1 `CellValue`

```typescript
/**
 * Discriminated union of every concrete value an XLSX cell may hold.
 *
 * Mapping to OOXML `<c t="…">`:
 *   number  ↔ t="n"
 *   string  ↔ t="s" (SST), t="str" (formula result), t="inlineStr"
 *   boolean ↔ t="b"
 *   null    ↔ <c/> (empty / blank cell, also used for unevaluated formulas)
 *   CellError ↔ t="e"
 *   Date    ↔ t="d" (ISO 8601 date), or t="n" with date-typed numFmt
 *   RichText ↔ t="s" with rich-text SST entry, or t="inlineStr" with <r> runs
 */
export type CellValue = number | string | boolean | null | CellError | DateCell | RichText;

/**
 * A date-typed value. Two storage shapes survive parse:
 *   - OOXML `t="d"` cells produce `{ kind: "date", iso }`.
 *   - OOXML `t="n"` cells with a date-detected `numFmt` produce
 *     `{ kind: "date", serial }` (number of days since 1900-01-00 or
 *     1904-01-01 depending on `XlsxWorkbook.date1904`).
 *
 * The model preserves whichever was on disk; the renderer and formula
 * engine convert at the boundary. Internal date arithmetic uses the
 * serial.
 */
export interface DateCell {
  readonly kind: "date";
  readonly serial?: number;
  readonly iso?: string;
}

/**
 * A rich-text value. Used for `t="s"` cells whose SST entry contains
 * `<r>` runs and for `t="inlineStr"` cells with rich runs. Plain-text
 * SST hits emit `string`, not `RichText`.
 */
export interface RichText {
  readonly kind: "rich-text";
  /** Concatenated plain-text projection (for search, formula evaluation). */
  readonly plain: string;
  readonly runs: ReadonlyArray<RichRun>;
}

export interface RichRun {
  readonly text: string;
  readonly font?: Font;
}
```

### 4.2 `CellError`

```typescript
/**
 * Excel error sentinels. Stored as a typed enum (not as numeric codes)
 * so model consumers and serialized JSON snapshots are self-describing.
 * The `#GETTING_DATA` and `#SPILL!` values are Excel-2010+ / Excel-365
 * extensions; we accept and round-trip them but do not author them.
 */
export const enum CellError {
  Ref = "#REF!",
  Value = "#VALUE!",
  Div0 = "#DIV/0!",
  Name = "#NAME?",
  NA = "#N/A",
  Null = "#NULL!",
  Num = "#NUM!",
  GettingData = "#GETTING_DATA",
  Spill = "#SPILL!",
}
```

> NOTE: The OOXML wire form uses the literal error string in the
> `<v>` child of `<c t="e">`. The numeric error codes (0x00 = `#NULL!`,
> 0x07 = `#DIV/0!`, etc.) are an XLSB / BIFF concern; we do not see
> them at the XLSX layer.

## 5. Formula

```typescript
export interface Formula {
  /**
   * The original formula text without the leading `=`. Always the
   * source of truth — even when `tokens` is populated, the formula
   * engine re-parses on demand if it doesn't match.
   */
  readonly text: string;
  /**
   * Cached parse result. Populated lazily by the formula engine; never
   * persisted; cleared whenever `text` is mutated. Schema is defined in
   * `formula-engine.md`.
   */
  readonly tokens?: ReadonlyArray<unknown>;
  /**
   * Shared-formula metadata. When set, `text` is the master expression
   * for the corner cell (`x = 0, y = 0`); follower cells set `x`/`y`
   * to the offset from the corner and inherit the master text. Mirrors
   * Univer's `si + x + y` model and OOXML's `<f t="shared" si="N">`.
   */
  readonly shared?: { readonly si: number; readonly x: number; readonly y: number };
  /**
   * Anchor for an array formula. Set on every cell within the range,
   * but `text` is non-empty only on the corner cell. Mirrors OOXML's
   * `<f t="array" ref="A1:B5">`.
   */
  readonly arrayRef?: RangeRef;
}
```

> NOTE: We do **not** model dynamic-array spill ranges (`<f t="dynamicArray">`)
> as first-class formulas in P0. The corner cell's formula is parsed as a
> normal expression and `Cell.dynamicArray` carries the spill flag; the
> follower cells in the spill region show their cached values without a
> formula reference. This is consistent with Excel < 2019 round-trip.

## 6. Style table

The XLSX style model keeps **two** parallel tables: the workbook-level
shared resources (`fonts`, `fills`, `borders`, `numFmts`, `dxfs`,
`tableStyles`) and the per-cell `cellXfs` index that combines them.
`Cell.styleId` is the index into `cellXfs`.

```typescript
export interface StyleTable {
  /**
   * Number-format definitions keyed by `numFmtId`. IDs 0-163 are
   * built-ins (we keep a const map of those in `model/numfmt-builtins.ts`);
   * IDs ≥ 164 are custom and live in the workbook's `<numFmts>`.
   */
  readonly numFmts: ReadonlyMap<number, string>;
  /** Indexed font records. Cell xfs reference these by index. */
  readonly fonts: ReadonlyArray<Font>;
  readonly fills: ReadonlyArray<Fill>;
  readonly borders: ReadonlyArray<Border>;
  /**
   * `cellStyleXfs` — the named-style table. Cell xfs may reference an
   * entry here via `xfId`. We model these as full `CellFormat`s for
   * fidelity. Many real workbooks use only the implicit "Normal" entry
   * at index 0.
   */
  readonly cellStyleXfs: ReadonlyArray<CellFormat>;
  /**
   * `cellXfs` — the per-cell style table. `Cell.styleId` is an index
   * into this array.
   */
  readonly cellXfs: ReadonlyArray<CellFormat>;
  /**
   * Differential formats — the per-rule overlays referenced by
   * `ConditionalFormatRule.dxfId` and table-style records. We do not
   * author dxfs in P0 but preserve them through round-trip.
   */
  readonly dxfs: ReadonlyArray<Partial<CellFormat>>;
  /**
   * Named cell styles ("Normal", "Heading 1", "Comma"). Each entry
   * references a `cellStyleXfs` index and an optional builtinId.
   */
  readonly namedStyles: ReadonlyArray<{
    readonly name: string;
    readonly xfId: number;
    readonly builtinId?: number;
  }>;
  /**
   * Table styles. Preserved opaquely in P0 — pivot/list-object styling
   * is out of scope for the agent.
   */
  readonly tableStyles: ReadonlyArray<OpaqueXml>;
}

export interface CellFormat {
  /** Index into `StyleTable.numFmts`. 0 = General. */
  readonly numFmtId?: number;
  /** Index into `StyleTable.fonts`. */
  readonly fontId?: number;
  /** Index into `StyleTable.fills`. */
  readonly fillId?: number;
  /** Index into `StyleTable.borders`. */
  readonly borderId?: number;
  /** Index into `StyleTable.cellStyleXfs` (the parent named-style xf). */
  readonly xfId?: number;
  readonly alignment?: Alignment;
  readonly protection?: { readonly locked?: boolean; readonly hidden?: boolean };
  /**
   * `applyX` flags — when `false`, the cell xf inherits the
   * corresponding aspect from its `cellStyleXfs` parent rather than
   * using its own. Preserved verbatim.
   */
  readonly applyNumberFormat?: boolean;
  readonly applyFont?: boolean;
  readonly applyFill?: boolean;
  readonly applyBorder?: boolean;
  readonly applyAlignment?: boolean;
  readonly applyProtection?: boolean;
  readonly quotePrefix?: boolean;
  readonly pivotButton?: boolean;
}
```

### 6.1 `Font`

```typescript
export interface Font {
  /** Typeface name (`<name val="…"/>`). */
  readonly name?: string;
  /** Size in points (`<sz val="…"/>`). */
  readonly size?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  /** `"single"`, `"double"`, `"singleAccounting"`, `"doubleAccounting"`. Boolean shorthand for "single". */
  readonly underline?: boolean | "single" | "double" | "singleAccounting" | "doubleAccounting";
  readonly strike?: boolean;
  /** Color: either an explicit RRGGBB hex or a theme reference. */
  readonly color?: Color;
  /** Vertical alignment for the font (super/subscript). */
  readonly vertAlign?: "baseline" | "superscript" | "subscript";
  /** Font family code (`<family val="…"/>`). */
  readonly family?: number;
  /** Theme font scheme reference (`<scheme val="major" | "minor"/>`). */
  readonly scheme?: "major" | "minor";
}
```

### 6.2 `Fill`

```typescript
export interface Fill {
  /** Fill kind. `"none"` and `"solid"` map to `<patternFill>` with the same `patternType`. */
  readonly type: "none" | "solid" | "pattern" | "gradient";
  /** Foreground color (`<fgColor>`). For solid fills this is the visible color. */
  readonly color?: Color;
  /** Background color (`<bgColor>`). For pattern fills, the pattern field. */
  readonly bgColor?: Color;
  /**
   * Pattern type code (e.g. `"darkHorizontal"`, `"gray125"`). Required
   * for `type === "pattern"`. Optional for `type === "solid"`
   * (defaults to `"solid"`).
   */
  readonly patternType?: string;
  /**
   * Gradient fill detail. We do not author gradients in P0; preserved
   * opaquely so round-trip is lossless.
   */
  readonly gradient?: OpaqueXml;
}
```

### 6.3 `Border`

```typescript
export interface Border {
  readonly top?: BorderEdge;
  readonly right?: BorderEdge;
  readonly bottom?: BorderEdge;
  readonly left?: BorderEdge;
  readonly diagonal?: BorderEdge;
  /** When `true`, the diagonal runs from bottom-left to top-right. */
  readonly diagonalUp?: boolean;
  /** When `true`, the diagonal runs from top-left to bottom-right. */
  readonly diagonalDown?: boolean;
}

export interface BorderEdge {
  /**
   * Border line style. The full OOXML enumeration is preserved; the
   * renderer maps unknown values to `"thin"`. Empty / missing edge =
   * no border on that side.
   */
  readonly style:
    | "none"
    | "thin"
    | "medium"
    | "dashed"
    | "dotted"
    | "thick"
    | "double"
    | "hair"
    | "mediumDashed"
    | "dashDot"
    | "mediumDashDot"
    | "dashDotDot"
    | "mediumDashDotDot"
    | "slantDashDot";
  readonly color?: Color;
}
```

### 6.4 `Alignment`

```typescript
export interface Alignment {
  /** `<alignment horizontal="…"/>`. */
  readonly horizontal?:
    | "general"
    | "left"
    | "center"
    | "right"
    | "fill"
    | "justify"
    | "centerContinuous"
    | "distributed";
  /** `<alignment vertical="…"/>`. */
  readonly vertical?: "top" | "center" | "bottom" | "justify" | "distributed";
  readonly wrapText?: boolean;
  /**
   * Rotation in degrees (-90 … 180; 255 = vertical text per the spec).
   * Stored as the raw OOXML integer.
   */
  readonly textRotation?: number;
  /** Indent level (0..15). */
  readonly indent?: number;
  readonly shrinkToFit?: boolean;
  /**
   * Reading order: 0 = context, 1 = left-to-right, 2 = right-to-left.
   * Preserved as the raw OOXML integer.
   */
  readonly readingOrder?: 0 | 1 | 2;
}
```

### 6.5 `Color`

```typescript
/**
 * A color is one of:
 *   - `{ rgb: "FF112233" }` — 8-hex ARGB form (high byte = alpha).
 *     OOXML always uses ARGB; AABBGGRR is not valid.
 *   - `{ theme: 4, tint: -0.25 }` — theme-color reference. `tint` is
 *     a float in [-1.0, 1.0]; -1 is full black, +1 is full white.
 *   - `{ indexed: 64 }` — legacy palette index (index 64 = system
 *     foreground, 65 = system background). Preserved verbatim.
 *   - `{ auto: true }` — Excel's `auto` color (system foreground).
 */
export type Color =
  | { readonly rgb: string }
  | { readonly theme: number; readonly tint?: number }
  | { readonly indexed: number }
  | { readonly auto: true };
```

## 7. Defined names

```typescript
export interface DefinedName {
  readonly name: string;
  /**
   * `"workbook"` for a workbook-level name; otherwise the 0-based
   * sheet index the name is local to (mirrors OOXML's
   * `localSheetId` attribute).
   */
  readonly scope: "workbook" | number;
  /**
   * The reference body. May be a single ref, a range, a list of
   * ranges (semicolon-joined in OOXML), or a formula expression.
   * Stored as the original string for fidelity; consumers parse with
   * `parseDefinedNameRef(name.ref)` from `model/refs.ts`.
   */
  readonly ref: string;
  readonly comment?: string;
  readonly hidden?: boolean;
  /**
   * `_xlnm._FilterDatabase` and other reserved names are flagged so the
   * agent surface refuses to mutate them.
   */
  readonly builtin?: boolean;
}
```

## 8. Comments

```typescript
export interface Comment {
  /** Stable id, minted by the parser. Never recycled. */
  readonly id: number;
  /** Anchor cell. Always single-cell in OOXML. */
  readonly ref: CellRef;
  /** Author display name. */
  readonly author: string;
  /** Comment body as a flat run sequence (preserves rich-text markup). */
  readonly text: ReadonlyArray<RichRun>;
  /**
   * Modern threaded-comments timestamp (ISO 8601). Absent on legacy
   * `xl/comments{N}.xml`-only comments.
   */
  readonly createdAt?: string;
  /**
   * For replies in a threaded comment. Points at the parent comment's
   * `id`. Top-level comments leave this undefined.
   */
  readonly threadParentId?: number;
  /** Resolved-state flag, threaded comments only. */
  readonly resolved?: boolean;
}
```

> NOTE: OOXML stores comments in two parallel parts when threading is in
> use: the legacy `xl/comments{N}.xml` (text + author + cell ref) and
> the modern `xl/threadedComments/threadedComment{N}.xml` (id, parent
> id, timestamp, resolved-state, person-id reference into
> `xl/persons/person.xml`). The model collapses both into one
> `Comment` per logical thread node. The serializer reconstitutes both
> parts from the same data.

## 9. Hyperlinks

```typescript
export interface Hyperlink {
  /** Stable id, minted by the parser. */
  readonly id: number;
  /** Anchor range — usually a single cell but may span a region. */
  readonly ref: RangeRef;
  /** What the hyperlink points at. Discriminated by `kind`. */
  readonly target: HyperlinkTarget;
  /** Tooltip shown on hover. */
  readonly tooltip?: string;
  /** Display text override; when absent, the cell's value renders. */
  readonly display?: string;
}

export type HyperlinkTarget =
  | { readonly kind: "url"; readonly url: string }
  | { readonly kind: "mailto"; readonly email: string; readonly subject?: string }
  /** Internal sheet reference, e.g. `Sheet1!A1` or `'My Sheet'!B5:C10`. */
  | { readonly kind: "sheet"; readonly ref: RangeRef }
  /** Workbook-scoped defined name reference. */
  | { readonly kind: "definedName"; readonly name: string }
  /** Anything we couldn't classify. The original `Target` string is preserved. */
  | { readonly kind: "raw"; readonly raw: string };
```

> NOTE: External hyperlinks (`url` / `mailto`) live in
> `xl/worksheets/_rels/sheet{N}.xml.rels` as `Target` attributes with
> `TargetMode="External"`; the `<hyperlink>` element in the sheet
> references them via `r:id`. Internal sheet hyperlinks use the
> `location` attribute on the `<hyperlink>` element directly and have
> no relationship entry. The parser resolves both into the same
> `HyperlinkTarget` union.

## 10. Merged cells

```typescript
/** A merged region on a sheet. Just a `RangeRef`; no per-merge state. */
export type MergedCell = RangeRef;
```

The "anchor" cell of a merge is always `(r1, c1)` and is the only one
that may carry a value or formula in OOXML. The other cells in the
merge are present in `cells` only when the source workbook explicitly
emitted them; the renderer ignores their values.

## 11. Conditional formatting

```typescript
/**
 * A single conditional-formatting rule. Discriminated by `kind`. The
 * `sqref` is the range(s) the rule applies to; the `dxfId` references
 * an entry in `StyleTable.dxfs` that supplies the visual overlay.
 *
 * Rules we model in P0 are listed below. Rules outside this set are
 * captured as `kind: "opaque"` and round-trip verbatim.
 */
export type ConditionalFormatRule =
  | CellValueRule
  | TextRule
  | DateOccurringRule
  | ColorScaleRule
  | DataBarRule
  | OpaqueRule;

interface RuleBase {
  /** Stable id minted by the parser. */
  readonly id: number;
  /** Application range(s). Multiple ranges separated by spaces in OOXML. */
  readonly sqref: ReadonlyArray<RangeRef>;
  /** 1-based priority — lower numbers win when rules overlap. */
  readonly priority: number;
  /** Whether this rule stops further rules from applying to its range. */
  readonly stopIfTrue?: boolean;
  /** dxf overlay (cellFormat) to apply when the rule fires. */
  readonly dxfId?: number;
}

export interface CellValueRule extends RuleBase {
  readonly kind: "cellValue";
  /** Comparison operator. */
  readonly operator:
    | "lessThan"
    | "lessThanOrEqual"
    | "equal"
    | "notEqual"
    | "greaterThanOrEqual"
    | "greaterThan"
    | "between"
    | "notBetween";
  /**
   * Operand expressions in OOXML formula syntax. `between` /
   * `notBetween` carry two; other operators carry one. Parsed by the
   * formula engine on demand.
   */
  readonly formulas: ReadonlyArray<string>;
}

export interface TextRule extends RuleBase {
  readonly kind: "textContains";
  readonly operator: "containsText" | "notContainsText" | "beginsWith" | "endsWith";
  readonly text: string;
}

export interface DateOccurringRule extends RuleBase {
  readonly kind: "dateOccurring";
  readonly timePeriod:
    | "today"
    | "yesterday"
    | "tomorrow"
    | "last7Days"
    | "thisWeek"
    | "lastWeek"
    | "nextWeek"
    | "thisMonth"
    | "lastMonth"
    | "nextMonth";
}

export interface ColorScaleRule extends RuleBase {
  readonly kind: "colorScale";
  /** Two- or three-color stop list, in OOXML order. */
  readonly stops: ReadonlyArray<{
    readonly type: "min" | "max" | "num" | "percent" | "percentile" | "formula";
    readonly value?: string;
    readonly color: Color;
  }>;
}

export interface DataBarRule extends RuleBase {
  readonly kind: "dataBar";
  readonly minLength?: number; // percent (default 10)
  readonly maxLength?: number; // percent (default 90)
  readonly color: Color;
  /** P0 does not author axis customization; preserve `<extLst>` opaquely. */
  readonly opaqueExt?: OpaqueXml;
}

export interface OpaqueRule extends RuleBase {
  readonly kind: "opaque";
  /** The `<cfRule>` element verbatim (used for icon sets, top10, etc.). */
  readonly raw: OpaqueXml;
}
```

> NOTE: OOXML wraps a list of `<cfRule>` elements inside one or more
> `<conditionalFormatting sqref="…">` parents. We flatten that on
> parse — every `ConditionalFormatRule` carries its own `sqref` so
> we don't need to track the parent grouping. On serialization we
> regroup by identical `sqref` for compactness.

## 12. References (`CellRef`, `RangeRef`)

```typescript
/**
 * A single-cell reference. The `sheet` field is set when the reference
 * crossed a sheet boundary (e.g. `Sheet2!A1`) and undefined when the
 * ref is implicitly local to its owning sheet. Sheet name is stored
 * unquoted; the formatter re-adds quotes when needed.
 */
export interface CellRef {
  readonly sheet?: string;
  readonly row: number; // 0-based
  readonly col: number; // 0-based
  readonly absRow: boolean;
  readonly absCol: boolean;
}

/**
 * A rectangular range. `r1/c1` is the top-left, `r2/c2` the
 * bottom-right (both inclusive, 0-based). For a single-cell range
 * `r1 === r2 && c1 === c2`. Whole-row references encode as
 * `c1 = 0, c2 = MAX_COL`; whole-column as `r1 = 0, r2 = MAX_ROW`.
 *
 * Each endpoint carries its own absoluteness flags so we round-trip
 * `=$A1:B$5` verbatim.
 */
export interface RangeRef {
  readonly sheet?: string;
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
  readonly absR1: boolean;
  readonly absC1: boolean;
  readonly absR2: boolean;
  readonly absC2: boolean;
}

/** Excel's hard limits — the upper bounds of a `RangeRef`. */
export const MAX_ROW = 1_048_575; // row 1,048,576 in 1-based / Excel-A1
export const MAX_COL = 16_383; // column XFD in A1
```

### 12.1 A1 / R1C1 conversion helpers

All functions live in `packages/xlsx/src/model/refs.ts` and are pure
(no DOM, no DI, no IO). They are the **only** sanctioned conversion
between the boundary form (1-based, A1 / `Sheet1!A1`) and the internal
form (0-based row/col indices).

```typescript
/** "B" → 1, "AA" → 26, "XFD" → 16383. Throws on empty or non-letter input. */
export function colLettersToIndex(letters: string): number;
/** 0 → "A", 25 → "Z", 26 → "AA". Throws on negative input. */
export function colIndexToLetters(col: number): string;

/**
 * "A1" → { row: 0, col: 0, absRow: false, absCol: false }.
 * Accepts `$A$1`, `$A1`, `A$1`. Does NOT accept a sheet prefix.
 */
export function a1ToRowCol(a1: string): { row: number; col: number; absRow: boolean; absCol: boolean };

/** Inverse of `a1ToRowCol`. `{ absRow: true, absCol: true }` → `$A$1`. */
export function rowColToA1(row: number, col: number, absRow?: boolean, absCol?: boolean): string;

/**
 * Parse a fully-qualified range reference, e.g. `"Sheet1!A1"`,
 * `"'My Sheet'!A1:B5"`, or `"$A$1:$B$5"` (no sheet prefix). Returns
 * `RangeRef`. Single-cell input is normalized to `r1 === r2`,
 * `c1 === c2`. Throws on malformed input via `RefParseError`.
 */
export function parseRangeRef(input: string): RangeRef;

/**
 * Format a `RangeRef` back to its A1 string form. The sheet name is
 * quoted if it contains spaces, single quotes, or any character outside
 * `[A-Za-z0-9_]`. Single-cell ranges format as `"A1"` not `"A1:A1"`.
 */
export function formatRangeRef(ref: RangeRef): string;

/**
 * A1 ↔ R1C1 utilities. R1C1 is a *display* convention; we never store
 * it. These helpers exist for the formula bar, which surfaces R1C1 to
 * the user when the workbook's `WorkbookView.ref` requests it.
 */
export function a1ToR1c1(a1: string, anchor: { row: number; col: number }): string;
export function r1c1ToA1(r1c1: string, anchor: { row: number; col: number }): string;
```

> NOTE: 1-based at the boundary, 0-based internally. The agent surface,
> the renderer, and the OOXML parser/serializer all speak A1 strings
> (e.g. `"B5:C10"` is rows 5-10, columns B-C, 1-based). All in-memory
> structs (`Cell.row`, `Sheet.cells` keys, `RangeRef.r1`, …) are
> 0-based. The conversion happens **exclusively** in `model/refs.ts`;
> code outside this module that performs `±1` arithmetic on a row or
> column is a bug.

### 12.2 Range adjustment on row/col mutation

When a row or column is inserted or deleted, every `RangeRef` and every
formula token must be adjusted. The shared mechanism is in
`packages/xlsx/src/model/range-adjust.ts`:

```typescript
export type StructuralEdit =
  | { readonly kind: "insert-rows"; readonly sheetId: string; readonly atRow: number; readonly count: number }
  | { readonly kind: "delete-rows"; readonly sheetId: string; readonly atRow: number; readonly count: number }
  | { readonly kind: "insert-cols"; readonly sheetId: string; readonly atCol: number; readonly count: number }
  | {
      readonly kind: "delete-cols";
      readonly sheetId: string;
      readonly atCol: number;
      readonly count: number;
    };

/**
 * Adjust a `RangeRef` for a structural edit. Returns the new range, or
 * `null` when the edit fully deleted the range (e.g. delete-rows that
 * covers the entire ref). The cross-sheet case is short-circuited:
 * a ref with a different `sheet` than the edit's `sheetId` returns
 * unchanged.
 */
export function adjustRangeForEdit(ref: RangeRef, edit: StructuralEdit): RangeRef | null;
```

The same function is invoked from formula-token visitors (`Formula.text`
gets re-emitted with adjusted refs) and from the index of every
feature that owns a range — `MergedCell[]`, `Sheet.hyperlinks`,
`Sheet.comments`, `Sheet.conditionalFormats`, and `DefinedName[]`.

## 13. Shared-strings table

```typescript
export interface SharedStringTable {
  /**
   * The flat list of SST entries in OOXML order. `Cell.value === string`
   * cells whose source was `<c t="s"><v>N</v></c>` resolve to
   * `entries[N].plain` (or `entries[N].rich` when the entry contained
   * `<r>` runs).
   *
   * Entries are append-only: indices never change once parsed. New
   * strings introduced by mutations append; old slots are never
   * reused even when their last reference is removed (consistent with
   * Excel behavior).
   */
  readonly entries: ReadonlyArray<SharedStringEntry>;
  /**
   * Reverse index: `plain` text → first matching entry index. Used by
   * the mutation layer to dedupe. Rich-text entries are NOT deduped via
   * this index; mutation handlers compare the full `RichText` shape.
   */
  readonly indexByPlain: ReadonlyMap<string, number>;
}

export interface SharedStringEntry {
  /** Concatenated plain-text projection. */
  readonly plain: string;
  /** Rich-text runs when the SST entry carried `<r>` children. */
  readonly rich?: RichText;
}
```

## 14. Opaque XML carrier

```typescript
/**
 * Same shape as `@officeai/docx`'s `OpaqueXml`. Carries a verbatim
 * `fast-xml-parser` `preserveOrder` subtree so the serializer can emit
 * an unmodelled element back at its original location, with original
 * attributes and namespace prefixes, byte-equivalent (modulo XML
 * canonicalization — see `serializer.md` §Whitespace).
 */
export interface OpaqueXml {
  readonly tag: string;
  readonly attrs: Readonly<Record<string, string>>;
  readonly subtree: ReadonlyArray<unknown>;
  readonly rawAttrs: Readonly<Record<string, string>>;
}
```

## 15. Identity rules

These mirror the identity rules from
[`spec/shared/document-model.md`](../shared/document-model.md) §Identity.

1. **Every structural node carries a `NodeId`.** The parser mints them
   via the shared `defaultIdMinter` from `@officeai/core`. Nodes:
   `XlsxWorkbook`, every `Sheet`, every `Cell`, every `Comment`, every
   `Hyperlink`, every `ConditionalFormatRule`, every `DefinedName`.
2. **Ids are stable across mutations.** A cell whose value changes
   keeps its `Cell.row/col` (which together act as its identity key —
   see §3.1) and any consumer that captured a reference still sees the
   updated cell. Sheet rename keeps `Sheet.id` and `Sheet.sheetId`;
   only `Sheet.name` changes.
3. **A node deleted by one mutation and re-inserted by a later
   mutation gets a new id.** Ids are not recycled. This is what makes
   diffs between revisions sound.
4. **Sheet `index` is positional.** When a sheet is reordered, every
   sheet's `index` may change — but its `id` and `sheetId` do not.
5. **Snapshots are immutable.** Mutations produce new `XlsxSnapshot`
   instances with bumped `revision`. The previous snapshot remains
   valid and addressable for diff purposes.
6. **Sheet `id` vs `sheetId`.** `Sheet.id` is our minted `NodeId`;
   `Sheet.sheetId` is the OOXML `sheetId` attribute (a number stored
   as string, used for relationship targeting). Both are stable; only
   `name` and `index` may change.

## 16. Position and Selection

```typescript
/** A single cell address used by the renderer's selection model. */
export interface XlsxPosition {
  /** Sheet `id` (NodeId), not name and not index. */
  readonly sheet: NodeId;
  readonly row: number; // 0-based
  readonly col: number; // 0-based
}

export interface XlsxSelection {
  /** Anchor cell — the one the user originally clicked. */
  readonly anchor: XlsxPosition;
  /**
   * Selected ranges. Multi-selection (Ctrl+click) is supported by
   * carrying multiple ranges; single-cell selections have one range
   * with `r1 === r2` and `c1 === c2`.
   */
  readonly ranges: ReadonlyArray<RangeRef>;
}
```

> NOTE: Selection is **not** part of the persisted model. It lives on
> the editor session and is communicated to the renderer via a separate
> channel. The shared `Position` / `Selection` types in
> `spec/shared/document-model.md` are intentionally `unknown` — every
> format binds them on its own.

## 17. What lives in the shared core vs what is xlsx-specific

| Concern                                                 | Lives in `@officeai/core`               | XLSX-specific (this package)                                |
| ------------------------------------------------------- | --------------------------------------- | ----------------------------------------------------------- |
| `DocumentSnapshot`, `DocumentDiff`, `DiffChange`        | yes                                     | no — we extend, not redefine                                |
| `NodeId` minting (`defaultIdMinter`, `IdMinter`)        | yes                                     | no — consumed                                               |
| Snapshot freezing (`freezeSnapshot`)                    | yes (dev/test)                          | no                                                          |
| SHA-256 (`sha256Hex`) for `partHashes`                  | yes                                     | no                                                          |
| `OoxmlContainer` (zip walking, part read/write/clone)   | yes                                     | no — consumed via `ooxml.OoxmlContainer`                    |
| `Relationship` shape (rels-graph)                       | yes                                     | re-exported as `XlsxRelationship` for ergonomics            |
| `OpaqueXml` carrier shape                               | duplicated per-format (small, no value) | yes — local copy with the same fields as the docx one       |
| `parseXml` / `serializeXml` (`fast-xml-parser` adapter) | yes                                     | no — consumed                                               |
| Workbook / Sheet / Cell types                           | no                                      | **yes** — defined here                                      |
| Style table, shared strings, defined names              | no                                      | **yes** — XLSX-only                                         |
| Formula AST and engine                                  | no                                      | **yes** — see `formula-engine.md`                           |
| Range / Cell ref types and helpers                      | no                                      | **yes** — `model/refs.ts`                                   |
| Conditional-formatting rule types                       | no                                      | **yes**                                                     |
| Comment / hyperlink types                               | no                                      | **yes** — even though docx has comments, the shapes diverge |

## 18. Invariants enforced by the parser

1. `Sheet.cells` is sparse — there is no entry for empty cells unless
   the source workbook explicitly emitted a stub (`<c r="A1"/>` with
   no value). When stubs were emitted, they parse to a `Cell` with
   `value: null` and no `formula`, `styleId`, `hyperlinkId`, or
   `commentId`.
2. `Sheet.dimensions` always covers every cell key in `Sheet.cells`.
   When `cells` is empty, `dimensions` is `{ r1: 0, c1: 0, r2: 0, c2: 0 }`
   (the renderer shows an empty A1 region).
3. Every `Cell.styleId` indexes a valid `cellXfs` entry; out-of-range
   ids are coerced to `0` and a warning is logged via
   `XlsxParseWarning("dangling-style-id", …)`.
4. Every `Cell.hyperlinkId` and `Cell.commentId` indexes a valid entry
   in the corresponding sheet array; dangling ids are dropped.
5. Every `MergedCell` is rectangular and disjoint from every other
   merge on the same sheet. Overlapping merges are reported via
   `XlsxParseWarning("overlapping-merge", …)` and the later one wins.
6. Every `DefinedName.scope` of type `number` references a valid sheet
   index. Out-of-range references promote to `"workbook"` scope and
   emit `XlsxParseWarning("dangling-defined-name-scope", …)`.
7. Every `Formula.shared` carries a `si` that matches at least one
   other cell's `Formula.shared.si` on the same sheet. Single-member
   shared groups are unshared on parse and re-emitted as standalone
   `<f>` elements.
8. Every `RangeRef` satisfies `r1 ≤ r2 && c1 ≤ c2` and lies within
   `[0, MAX_ROW] × [0, MAX_COL]`. Out-of-bounds refs become
   `Cell.value === CellError.Ref` for the formula they appeared in.
9. `partHashes` covers every entry in `OoxmlContainer.parts` keyed by
   full zip path. The serializer's byte-preservation invariant is:
   `for each path P, partHashes[P] === sha256(container.parts.get(P).bytes)`
   for every untouched part.

## 19. Open questions / deferrals (P1)

- Modeled gradient fills: today `Fill.gradient` is `OpaqueXml`. Promoting
  it to a typed `GradientFill` is in P1 alongside the conditional-format
  data-bar / icon-set authoring story.
- Pivot-table-aware `ConditionalFormatRule`s: the OOXML grammar
  permits CF rules to reference pivot fields. We capture them as
  `kind: "opaque"` in P0; pivot-aware parsing follows the pivot model
  (which is itself out of scope per `feature-scope.md`).
- Per-cell `r:id` resolution beyond hyperlinks. Currently any unmodelled
  `<c>` attribute (`cm`, `ph`, `vm`) is preserved by the SheetJS adapter
  in the cell's pass-through projection but is not surfaced on the
  typed `Cell`. P1 promotes `cm` (dynamic-array) to a typed flag (already
  done) and `ph` (phonetic) to an opaque carrier on the cell.
