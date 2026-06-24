# XLSX — Pivot tables

> Status: F1. Promotes `xl/pivotCache/*` and `xl/pivotTables/*`
> from opaque preservation to a typed, editable model. Phased
> landing: Phase 1 (render) + Phase 2 (refresh) + Phase 3
> (create + basic edit) ship in this session; Phase 4 (slicers,
> calculated fields, OLAP/CUBE) is specified and stubbed for
> follow-on shifts. Supersedes the "Pivot tables — out of scope"
> entry in [`feature-scope.md`](feature-scope.md).

## Why

Pivot tables are the #1 reason analysts open Excel. Today our
`xlsx:` commands and the grid renderer treat them as opaque XML
blobs that round-trip but render as static cell values frozen at
the time the source workbook was last saved. Editing the source
data leaves the pivot stale. There is no way to create a new
pivot from the UI. Spreadsheet users expect this to work inside any
credible document workspace.

## OOXML mapping

```
xl/pivotCache/
  cacheDefinitionN.xml      ← schema (fields, source range, refresh metadata)
  cacheRecordsN.xml         ← raw records (one row per source row)
xl/pivotTables/
  pivotTableN.xml           ← layout (rows / cols / data / filter / page axes)
  _rels/pivotTableN.xml.rels  ← link to the cache definition
xl/_rels/workbook.xml.rels    ← workbook ↔ caches
xl/worksheets/sheetK.xml      ← `<pivotTableParts>` → tables for that sheet
xl/styles.xml                 ← `<pivotStyles>` (we already preserve this)
[Content_Types].xml           ← overrides for new caches/tables
```

Relationship types:

| Type URI suffix                       | From                | To                                   |
| ------------------------------------- | ------------------- | ------------------------------------ |
| `/relationships/pivotCacheDefinition` | `xl/workbook.xml`   | `xl/pivotCache/cacheDefinitionN.xml` |
| `/relationships/pivotCacheRecords`    | cache definition    | `xl/pivotCache/cacheRecordsN.xml`    |
| `/relationships/pivotTable`           | `xl/worksheets/...` | `xl/pivotTables/pivotTableN.xml`     |

`workbook.xml` carries `<pivotCaches><pivotCache cacheId="..." r:id="..."/></pivotCaches>`.
We extend our existing typed workbook to model `pivotCaches`
explicitly (today preserved on the part).

## Typed model

```ts
// packages/xlsx/src/model/pivot.ts

export type PivotAggregation =
  | "sum"
  | "count"
  | "average"
  | "max"
  | "min"
  | "product"
  | "countNums"
  | "stdDev"
  | "stdDevp"
  | "var"
  | "varp";

export type PivotShowAs =
  | { kind: "normal" }
  | { kind: "percentOfTotal" }
  | { kind: "percentOfCol" }
  | { kind: "percentOfRow" }
  | { kind: "difference"; baseField: string; baseItem: string }
  | { kind: "percentOfDifference"; baseField: string; baseItem: string }
  | { kind: "runningTotal"; baseField: string }
  | { kind: "rank"; baseField: string };

export interface PivotCacheField {
  readonly name: string;
  readonly dataType: "string" | "number" | "boolean" | "date" | "mixed";
  /** Discrete shared items table (`<sharedItems>`) — also used for column ordering. */
  readonly items: ReadonlyArray<string | number | boolean | null>;
  /** Optional grouping (date or numeric ranges). */
  readonly grouping?: PivotFieldGrouping;
  readonly raw?: Readonly<Record<string, unknown>>;
}

export type PivotFieldGrouping =
  | { kind: "range"; start: number; end: number; interval: number }
  | { kind: "date"; by: ReadonlyArray<"years" | "quarters" | "months" | "days" | "hours" | "minutes"> }
  | { kind: "discrete"; mappedItemIndex: number };

export type PivotSource =
  | { kind: "worksheet"; sheet: string; range: string /* A1 */ }
  | { kind: "named"; name: string }
  | { kind: "external"; raw: string /* preserved verbatim */ };

export interface PivotCache {
  readonly id: number;
  readonly source: PivotSource;
  readonly fields: ReadonlyArray<PivotCacheField>;
  /** Materialised rows. For very large caches this is a streaming view. */
  readonly records: ReadonlyArray<ReadonlyArray<unknown>>;
  /** Last-refresh metadata (date, by, version) — preserved verbatim. */
  readonly refresh?: PivotCacheRefreshMeta;
  readonly raw?: Readonly<Record<string, unknown>>;
}

export interface PivotAxisField {
  /** Index into `cache.fields`. */
  readonly fieldIndex: number;
  /** Subtotal toggles (default true at axis-end, false at intermediate levels). */
  readonly subtotals?: ReadonlyArray<PivotAggregation>;
  /** Item-level filter (cache field item indexes hidden in this layout). */
  readonly hiddenItems?: ReadonlyArray<number>;
  /** Collapsed item indexes (UI-driven; persists in pivot XML). */
  readonly collapsedItems?: ReadonlyArray<number>;
}

export interface PivotDataField {
  readonly fieldIndex: number;
  readonly aggregation: PivotAggregation;
  readonly showAs: PivotShowAs;
  /** Number format id or custom format. */
  readonly numberFormat?: string;
  /** Display name (overrides "Sum of X"). */
  readonly displayName?: string;
}

export interface PivotFilterField {
  readonly fieldIndex: number;
  /** Selected cache item indexes (empty = all). */
  readonly selectedItems?: ReadonlyArray<number>;
}

export interface PivotTable {
  readonly id: number;
  readonly name: string;
  readonly cacheId: number;
  readonly sheet: string;
  /** Anchor: top-left cell of the rendered region (e.g. "A3"). */
  readonly anchor: string;
  readonly rowAxis: ReadonlyArray<PivotAxisField>;
  readonly colAxis: ReadonlyArray<PivotAxisField>;
  readonly dataAxis: ReadonlyArray<PivotDataField>;
  readonly filterAxis: ReadonlyArray<PivotFilterField>;
  readonly layout: PivotLayoutOptions;
  readonly styleName?: string;
  readonly raw?: Readonly<Record<string, unknown>>;
}

export interface PivotLayoutOptions {
  readonly compact: boolean; // compact form (default Excel 2007+)
  readonly outline: boolean; // outline form
  readonly tabular: boolean; // tabular form
  readonly rowGrandTotals: boolean;
  readonly colGrandTotals: boolean;
  readonly multipleFieldFilters: boolean;
  readonly mergeItem: boolean;
}
```

`XlsxWorkbook` gains:

```ts
readonly pivotCaches: ReadonlyArray<PivotCache>;
readonly pivotTables: ReadonlyArray<PivotTable>;
```

## Parser

`packages/xlsx/src/parser/pivot/cache.ts` reads each
`xl/pivotCache/cacheDefinitionN.xml` + matching records; honours
the rels graph for the records link. Walks `<sharedItems>` for
field values; reconstructs `records` from `<r>` rows.

`packages/xlsx/src/parser/pivot/table.ts` reads pivot table
parts, links by `cacheId`, materialises the axis arrays.

Add `xl/pivotCache/` and `xl/pivotTables/` to `MODELED_PREFIXES`
in `parse.ts`. Files we don't model (rare extensions) stay in
`opaqueParts` per the existing exclusion rules.

## Serializer

Mirror under `packages/xlsx/src/serializer/pivot/`. Round-trip
discipline:

- For pivots **untouched** by typed commands: re-emit from `raw`.
- For pivots **edited**: rebuild deterministically from typed
  fields; preserve any unmodelled `raw` attributes by merging
  them onto the rebuilt root element (same pattern as PPTX
  `spPrTail`).
- Update `[Content_Types].xml` overrides + workbook
  `<pivotCaches>` + sheet `<pivotTableParts>` graph.

## Phase 1 — render in grid

Pivot tables today render as cached cell values written into the
sheet XML at save time. We intercept the grid render pass:

1. After the grid resolves the static cell view for the visible
   region of a sheet, compile the pivot tree
   (`packages/xlsx/src/pivot/compile.ts`) for any pivot table
   anchored within that region.
2. The compiled tree emits a list of `(row, col, value, styleId,
group, level, isHeader, isTotal)` tuples that the renderer
   overlays on top of the static cells. Underlying sheet XML is
   unchanged.
3. Group rows render with a chevron disclosure (▶/▼) bound to
   `pivotField.items[i].sd` ("show details").

Style ids resolve through the workbook's existing
`<pivotStyles>` table (already preserved in `xl/styles.xml`).

## Phase 2 — refresh

`packages/xlsx/src/pivot/recompute.ts` pure function:

```ts
export function refreshPivotCache(
  cache: PivotCache,
  workbook: XlsxWorkbook,
  options?: { invalidateRecords?: boolean }
): { cache: PivotCache; affectedTables: ReadonlyArray<number> };
```

- Walks `cache.source` against the in-memory model (re-uses the
  sync formula engine for range resolution).
- Rebuilds `records[][]` row by row.
- Re-derives `<sharedItems>` per field, including type
  inference.
- Marks all dependent `PivotTable` ids as needing recompile.

Trigger:

- Explicit `xlsx:refresh-pivot { tableId | cacheId }` command.
- Optional auto-refresh on source range mutation
  (`workbook.options.autoRefreshPivots = true`, default `false`
  to preserve "manual" Excel behaviour).

## Phase 3 — create + basic edit

Commands (`packages/xlsx/src/commands/pivot-*.ts`):

| Command                           | Payload                                                         | Effect                                                 |
| --------------------------------- | --------------------------------------------------------------- | ------------------------------------------------------ |
| `xlsx:create-pivot-table`         | `{ source: PivotSource; anchorSheet; anchorCell; name? }`       | Mints cache + table, writes parts + rels + ContentType |
| `xlsx:add-pivot-field`            | `{ tableId; fieldIndex; axis: "row"\|"col"\|"data"\|"filter" }` | Appends to chosen axis                                 |
| `xlsx:remove-pivot-field`         | `{ tableId; axis; position }`                                   | Drops from axis                                        |
| `xlsx:move-pivot-field`           | `{ tableId; from: { axis, position }; to: { axis, position } }` | Reorders / reassigns                                   |
| `xlsx:set-pivot-data-aggregation` | `{ tableId; position; aggregation; showAs?; numberFormat? }`    | Mutates a `PivotDataField`                             |
| `xlsx:set-pivot-layout`           | `{ tableId; layout: Partial<PivotLayoutOptions> }`              | Compact / outline / tabular toggles + grand totals     |
| `xlsx:set-pivot-collapse`         | `{ tableId; axis; position; itemIndex; collapsed: boolean }`    | Drives the chevron in Phase 1                          |
| `xlsx:refresh-pivot`              | `{ tableId? }`                                                  | Phase-2 hook                                           |
| `xlsx:delete-pivot-table`         | `{ tableId }`                                                   | Removes table + cache (if last consumer) + rels        |

UI: a right-rail `PivotPanel.tsx` similar to Excel's:

```
┌ FIELDS ─────────────┐    ┌ FILTERS ────┐ ┌ COLUMNS ────┐
│ ☐ Region            │    │ Region      │ │ Quarter     │
│ ☐ Quarter           │    └─────────────┘ └─────────────┘
│ ☐ Product           │    ┌ ROWS ───────┐ ┌ VALUES ─────┐
│ ☐ Sales             │    │ Product     │ │ Sum of Sales│
│ ☐ Units             │    └─────────────┘ │ Avg Units   │
└─────────────────────┘                    └─────────────┘
```

Drag from FIELDS into a target zone → `xlsx:add-pivot-field`.
Drag between zones → `xlsx:move-pivot-field`. Click a value
chip → opens an aggregation picker (`xlsx:set-pivot-data-aggregation`).

## Phase 4 — specified, stubbed, not built tonight

### `GETPIVOTDATA` and CUBE functions

`packages/xlsx/src/formula/functions/getpivotdata.ts` and
`cube.ts` register stubs that return `#NAME?` until implemented.
Spec:

- `GETPIVOTDATA(data_field, pivot_table, [field1, item1], …)`
  resolves `pivot_table` to the anchor cell of a `PivotTable`,
  walks the compiled tree for the matching `(field, item)`
  intersection, and returns the data field's aggregated value.
- `CUBEMEMBER` / `CUBEVALUE` / `CUBESET` require an OLAP cube
  connection; we stub at the function level and reject any call
  in this phase.

### Slicers and timelines

`xl/slicers/`, `xl/slicerCaches/`, drawing anchors with
`<sle:slicer>` references.

- Phase 4 promotes them to typed `Slicer { id, sourcePivotIds,
filterFieldIndex, selection }` and `Timeline { id,
sourcePivotIds, dateFieldIndex, range }`.
- For Phase 1-3, slicers stay opaque and the
  `serializer/pivot/index.ts` ensures editing a typed pivot
  doesn't break the slicer rels graph (specific test:
  `pivot-typed-edit-preserves-slicer-roundtrip.test.ts`).

### Calculated fields / items

`<calculatedField>` / `<calculatedItem>` carry formula strings.
Phase 4 adds `PivotCalculatedField { name, formula }` to
`PivotTable`, evaluates via the existing formula engine on
refresh.

## Round-trip invariants

1. **No-edit byte-equivalence.** Loading and saving without
   touching pivots leaves every pivot part byte-identical.
2. **Single-table edit.** Editing layout on table N changes only
   `xl/pivotTables/pivotTableN.xml`; cache parts and other
   tables are byte-identical.
3. **Refresh changes records.** `xlsx:refresh-pivot` rewrites
   `cacheRecordsN.xml` deterministically; the unedited
   `cacheDefinitionN.xml` extension attributes survive.
4. **New pivot opens in Excel.** A pivot created via
   `xlsx:create-pivot-table` opens in Excel 2024 + LibreOffice
   without "needs repair" prompts. Validated by a LibreOffice
   roundtrip headless test.

## Acceptance criteria

A1. **Load.** `fixtures/xlsx/real-excel-mac-2021-pivot.xlsx`
parses into a non-empty `pivotCaches` and `pivotTables`.

A2. **Render.** The XLSX editor displays the pivot's compiled
cells (header / row / value / total) over the sheet at the
anchor.

A3. **Refresh.** Editing a source cell + dispatching
`xlsx:refresh-pivot` updates aggregated values.

A4. **Create.** `xlsx:create-pivot-table` over a 4-column 50-row
range produces a working pivot rendered at the anchor; saving and
reopening in LibreOffice shows the same pivot.

A5. **Phase 4 stub.** `GETPIVOTDATA` returns `#NAME?` with a
spec-referenced `unsupported` flag; the cell is not flagged as a
parse error.

## Out of scope for this session

- Slicers / timelines authoring (Phase 4).
- Calculated fields / items (Phase 4).
- `GETPIVOTDATA` / CUBE function evaluation (Phase 4 stubs).
- Power Pivot data model integration.
- Pivot charts (separate from pivot tables; deferred).
- Drag-on-canvas pivot anchor moves (use commands).
