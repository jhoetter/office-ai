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
  /**
   * Set when typed pivot edits land (Phase 3+). Phase 1 leaves this
   * unset / `false` so the serializer re-emits pivot parts
   * byte-identical from {@link PivotTablePart.raw} /
   * {@link PivotCachePart.raw}. Optional so existing call sites
   * that build a `XlsxDirtyFlags` literal don't need to grow a
   * field for a Phase-3 hook they can't trigger yet.
   */
  pivotTables?: boolean;
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
  readonly sheetjs: import("@e965/xlsx").WorkBook;
  /**
   * Image media keyed by `xl/media/...` part path. Sheet-level image
   * placements reference into this map via `SheetImage.mediaRef`, so
   * identical bytes uploaded twice round-trip through a single media
   * part (deduped on `ImageBlob.hash`).
   */
  readonly images: ReadonlyMap<string, ImageBlob>;
  /**
   * Workbook-scoped (or sheet-scoped) named ranges. Mirror of OOXML
   * `<workbook><definedNames><definedName name="…" localSheetId="…">
   * Sheet1!$A$1:$B$2</definedName></definedNames></workbook>`.
   *
   * Spec: see C12 in the night-shift unification plan.
   */
  readonly definedNames: ReadonlyArray<DefinedName>;
  /**
   * Pivot table parts (`xl/pivotTables/pivotTableN.xml`). Phase 1 of
   * `spec/xlsx/pivot-tables.md` keeps the entire pivot definition
   * opaque (`raw` carries the verbatim XML); only `name` is lifted
   * out of the root element so consumers can identify the table by
   * its display name. Round-trips byte-identical.
   */
  readonly pivotTables: ReadonlyArray<PivotTablePart>;
  /**
   * Pivot cache parts (`xl/pivotCache/pivotCacheDefinitionN.xml`)
   * with their matching `pivotCacheRecordsN.xml` bytes attached.
   * Same opaque-preservation contract as {@link XlsxWorkbook.pivotTables}.
   */
  readonly pivotCaches: ReadonlyArray<PivotCachePart>;
}

/**
 * Phase 1 typed slot for an XLSX pivot table part. The full
 * `<pivotTableDefinition>` element travels in {@link PivotTablePart.raw}
 * verbatim so it round-trips byte-identical even though we don't yet
 * model the rows / cols / data axes as typed structures. Phase 3
 * upgrades this to a full `PivotTable` (see
 * `spec/xlsx/pivot-tables.md`).
 */
export interface PivotTablePart {
  /** Container path, e.g. `xl/pivotTables/pivotTable1.xml`. */
  readonly partPath: string;
  /** Pivot table display name (`<pivotTableDefinition name="…">`). */
  readonly name: string;
  /**
   * `cacheId` attribute on the root element when present. Pivot
   * tables reference their cache by id (the cache also exists in
   * {@link XlsxWorkbook.pivotCaches}); we lift the id out so callers
   * can join the two sides without re-parsing the XML.
   */
  readonly cacheId?: number;
  /**
   * Stable OOXML id of the sheet that anchors this pivot
   * (matches {@link Sheet.sheetId}). Lifted at parse time from the
   * sheet whose rels reference this part — the pivot definition
   * itself never names its own anchor sheet, so the rels graph is
   * the only source of truth. `undefined` when the table is
   * defined in the package but not referenced from any worksheet
   * (rare, but legal in OOXML).
   */
  readonly sheetId?: string;
  /**
   * Geometry of the pivot's output rectangle on the anchor sheet,
   * lifted from `<location ref="A8:B12" firstHeaderRow="…" …/>`.
   * Phase 1 surfaces this so the grid can paint the pivot bounds
   * read-only without round-walking the raw XML on every render.
   * `undefined` when the part lacks a `<location>` element (a defect
   * in OOXML terms — Excel always writes one — but we degrade
   * gracefully rather than refusing to load).
   */
  readonly location?: PivotTableLocation;
  /** Verbatim part XML, ready for byte-identical re-emission. */
  readonly raw: string;
  /** Content type from `[Content_Types].xml` when present. */
  readonly contentType?: string;
  /**
   * Verbatim rels-part XML
   * (`xl/pivotTables/_rels/pivotTableN.xml.rels`), if any. The pivot
   * table's link to its cache definition lives there.
   */
  readonly relsXml?: string;
}

/**
 * Read-only geometry of a pivot table's output rectangle on its
 * anchor sheet. Mirrors the attributes Excel writes on
 * `<pivotTableDefinition><location/>`. All indices are 0-based and
 * inclusive at both ends, matching {@link Sheet.cells} keys.
 */
export interface PivotTableLocation {
  /** Original A1 string from `<location ref="A8:B12"/>`. */
  readonly ref: string;
  /** Inclusive 0-based bounds derived from `ref`. */
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
  /** Total rows / cols spanned (= r2-r1+1, c2-c1+1). */
  readonly rowCount: number;
  readonly colCount: number;
  /**
   * Row offset of the first header row, relative to `r1`. Almost
   * always `0` in Excel-emitted pivots. Lifted directly from the
   * `firstHeaderRow` attribute when present; defaults to `0`.
   */
  readonly firstHeaderRow: number;
  /**
   * Row offset of the first data row (i.e. first non-header row)
   * relative to `r1`. Lifted from `firstDataRow`.
   */
  readonly firstDataRow: number;
  /**
   * Column offset of the first data column relative to `c1`.
   * Lifted from `firstDataCol`. Row-label columns sit to the left
   * of this offset.
   */
  readonly firstDataCol: number;
}

/**
 * Phase 1 typed slot for an XLSX pivot cache. Pairs the
 * `pivotCacheDefinitionN.xml` part (`raw` + `partPath`) with the
 * sibling `pivotCacheRecordsN.xml` bytes when present. Both blobs
 * are opaque in Phase 1.
 */
export interface PivotCachePart {
  /** Container path of the cache definition. */
  readonly partPath: string;
  /**
   * Cache id matching the workbook's
   * `<pivotCaches><pivotCache cacheId="…" r:id="…"/></pivotCaches>`
   * entry. Lifted from the root element when present; falls back to
   * the index of the rels relationship when the cache file omits it
   * (legacy authoring tools occasionally do this).
   */
  readonly cacheId: number;
  /** Verbatim cache-definition XML. */
  readonly raw: string;
  readonly contentType?: string;
  readonly relsXml?: string;
  /**
   * Container path of the matching `pivotCacheRecordsN.xml` part
   * when the cache definition references one via its rels.
   */
  readonly recordsPartPath?: string;
  /** Verbatim records-part XML. */
  readonly recordsRaw?: string;
  readonly recordsContentType?: string;
}

/**
 * Workbook- or sheet-scoped named range (a.k.a. defined name).
 * `refersTo` is the raw OOXML reference text, with `=` prefix
 * stripped at parse time. We keep the original string verbatim so
 * round-trip survives even when the reference can't be parsed
 * cleanly (e.g. cross-sheet, cross-workbook, formulas).
 */
export interface DefinedName {
  /** Unique identifier within the snapshot (synthetic). */
  readonly id: string;
  /** Display name as used in formulas (`MyName`). Case-sensitive. */
  readonly name: string;
  /** OOXML refersTo text (without leading `=`). */
  readonly refersTo: string;
  /**
   * Sheet scope. `undefined` = workbook-scoped (visible from any
   * sheet); a string = scoped to the sheet with that name. OOXML
   * stores this as the 0-based `localSheetId` index — we resolve
   * it to a name at parse time so add / rename / delete are
   * straightforward.
   */
  readonly scope?: string;
  /** Optional comment shown in the Name Manager. */
  readonly comment?: string;
  /**
   * Hide the name from the Name Box dropdown. Excel uses this for
   * print areas and pivot caches; we preserve it but our UI never
   * sets it.
   */
  readonly hidden?: boolean;
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
   * 0-based column indices currently hidden via `<col hidden="1"/>`
   * in the source `<cols>` band, mirroring {@link Sheet.hiddenRows}.
   * The grid renders these columns at width 0 and the serializer is
   * expected to keep them hidden through the round-trip via the
   * captured `colsXml`. Mutating this set is not (yet) wired to a
   * command — the field exists so opening a workbook with hidden
   * columns doesn't surface them as full-width data on screen.
   */
  readonly hiddenCols: ReadonlySet<number>;
  /**
   * Per-sheet default column width in CSS pixels, derived from
   * `<sheetFormatPr defaultColWidth="..."/>`. When a column has no
   * explicit `<col>` entry and no `xlsx:set-column-width` override,
   * the renderer falls back to this (instead of the Grid's hard-coded
   * default) so workbooks with non-standard defaults look right at
   * open time. `undefined` = use the renderer's built-in default.
   */
  readonly defaultColWidthPx?: number;
  /**
   * Per-sheet default row height in CSS pixels, derived from
   * `<sheetFormatPr defaultRowHeight="..."/>`. Same fallback rules as
   * {@link Sheet.defaultColWidthPx}.
   */
  readonly defaultRowHeightPx?: number;
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
   * Frozen-pane configuration. `rows` rows from the top stay visible
   * when scrolling vertically; `cols` columns from the left stay
   * visible when scrolling horizontally. `0` on either axis means
   * "no freeze on this axis"; `undefined` means "no freeze at all".
   *
   * OOXML mirror: `<sheetView><pane xSplit="cols" ySplit="rows"
   * topLeftCell="…" state="frozen"/></sheetView>`. We only emit the
   * `frozen` state — Excel's `frozenSplit` (movable split bar) is a
   * niche feature we don't expose in the toolbar, so the
   * round-trip preserves it as opaque if a file already had it
   * (the parser falls back to ignoring non-`frozen` panes).
   */
  readonly freeze?: FreezePanes;
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
  /**
   * Typed conditional-formatting rules authored in this session.
   * Rules are evaluated at render time by the editor and applied
   * as visual overlays on top of the cell's base style.
   *
   * P0 round-trip: typed rules are NOT yet emitted to OOXML on
   * serialize — they live only in the in-memory model. Existing
   * `<conditionalFormatting>` blocks from the parsed workbook are
   * preserved verbatim via {@link Sheet.opaqueConditionalFormats}
   * so opening + saving doesn't destroy them.
   *
   * Spec: see C10 in the night-shift unification plan.
   */
  readonly conditionalFormats: ReadonlyArray<ConditionalFormat>;
  /**
   * Verbatim `<conditionalFormatting …>…</conditionalFormatting>`
   * blocks captured from the parsed worksheet. Re-emitted byte-
   * identically on dirty serialize so existing CF rules survive
   * round-trip even though we don't (yet) model them typed.
   */
  readonly opaqueConditionalFormats: ReadonlyArray<string>;
  /**
   * Typed data-validation rules authored in this session.
   * The editor surfaces these as in-cell dropdown arrows and
   * (optionally) input gating. Each rule covers one A1 range.
   *
   * P0 round-trip: typed `list` rules ARE emitted on serialize
   * (a new `<dataValidations>` block replaces any pre-existing
   * one for dirty sheets); other kinds are deferred to a future
   * pass and stay only in {@link Sheet.opaqueDataValidations}.
   *
   * Spec: see C11 in the night-shift unification plan.
   */
  readonly dataValidations: ReadonlyArray<DataValidation>;
  /**
   * Verbatim `<dataValidations …>…</dataValidations>` block
   * captured from the parsed worksheet (a single block per
   * sheet per OOXML spec). Re-emitted byte-identically on dirty
   * serialize when the typed list is empty so existing
   * non-`list` validations survive round-trip even though we
   * don't (yet) model them typed. When the user adds typed
   * `list` rules, the typed emitter takes over and the opaque
   * block is dropped — Excel's dialog would do the same.
   */
  readonly opaqueDataValidations?: string;
  /**
   * Excel Tables (a.k.a. ListObjects) attached to this sheet (C14).
   * Each entry maps 1:1 to an `xl/tables/tableN.xml` part. The
   * parser hydrates them, the serializer re-emits any tables flagged
   * dirty, and the editor renders alternating row banding + a bold
   * header band over `range` so the user gets the Excel-Tables look
   * even before structured references / auto-extend ship.
   */
  readonly tables: ReadonlyArray<TableDef>;
  /**
   * Charts attached to this sheet (C15). Each entry is a free-floating
   * overlay with a typed series binding so the editor can paint a
   * lightweight SVG preview without depending on a charting library.
   *
   * Round-trip caveat: existing chart parts (`xl/charts/*.xml`) live
   * in `opaqueParts` and survive saves verbatim. Charts added via
   * `xlsx:add-chart` are in-memory only — re-emitting brand-new
   * DrawingML chart parts ships in a follow-up.
   */
  readonly charts: ReadonlyArray<SheetChart>;
  /**
   * Verbatim `<hyperlinks>…</hyperlinks>` block captured from the
   * parsed worksheet XML. SheetJS does not surface sheet-level
   * hyperlinks, and a dirty sheet rewrite would otherwise drop the
   * block entirely. Re-injected by the serializer when the sheet is
   * dirty so URL / anchor hyperlinks survive every save. The matching
   * `r:id` rels live in the sheet rels part (preserved by the rels
   * round-trip path independently).
   */
  readonly hyperlinksXml?: string;
  /**
   * Verbatim `<tableParts>…</tableParts>` block captured from the
   * parsed worksheet XML. We re-emit it from the typed `tables`
   * array on dirty save so existing `<tableParts>` references
   * survive even when SheetJS regenerates the worksheet body.
   * Falls back to the original opaque block when typed `tables`
   * is empty (defensive).
   */
  readonly tablePartsXml?: string;
  /**
   * Verbatim `<cols>…</cols>` block captured from the parsed
   * worksheet XML. Re-injected on dirty save so source column
   * widths, hidden cols, custom widths, and outline levels survive
   * a sheet rewrite. When `Sheet.columnWidths` carries explicit
   * overrides we still re-inject the opaque block; future passes
   * will merge typed widths into it.
   */
  readonly colsXml?: string;
  /**
   * Verbatim `<sheetViews>…</sheetViews>` block captured from the
   * parsed worksheet XML. Re-injected on dirty save so zoom,
   * selection, gridline toggles, view mode, etc. all survive.
   * When the typed `freeze` field has been mutated, the serializer
   * surgically replaces just the `<pane>` element inside this
   * block instead of dropping it wholesale.
   */
  readonly sheetViewsXml?: string;
  /**
   * Verbatim `<sheetProtection …/>` element captured from the
   * parsed worksheet XML, when present. We do not author sheet
   * protection in this milestone — the value round-trips opaquely
   * so password-protected sheets do not silently lose protection
   * after a non-protection edit.
   */
  readonly sheetProtectionXml?: string;
  /**
   * Verbatim `<pageMargins …/>` element. Same opaque round-trip
   * pattern as `sheetProtectionXml`.
   */
  readonly pageMarginsXml?: string;
  /**
   * Verbatim `<pageSetup …/>` element. Page orientation, paper
   * size, scaling, etc. — preserved opaquely.
   */
  readonly pageSetupXml?: string;
  /**
   * Verbatim `<printOptions …/>` element. Preserved opaquely.
   */
  readonly printOptionsXml?: string;
  /** Verbatim `<headerFooter …>…</headerFooter>` block. */
  readonly headerFooterXml?: string;
  /** Verbatim `<rowBreaks …>…</rowBreaks>` block. */
  readonly rowBreaksXml?: string;
  /** Verbatim `<colBreaks …>…</colBreaks>` block. */
  readonly colBreaksXml?: string;
  /** Verbatim `<ignoredErrors …>…</ignoredErrors>` block. */
  readonly ignoredErrorsXml?: string;
  /** Verbatim `<legacyDrawing r:id="…"/>` element (anchors VML notes). */
  readonly legacyDrawingXml?: string;
  /** Verbatim `<legacyDrawingHF r:id="…"/>` element (header/footer VML). */
  readonly legacyDrawingHFXml?: string;
  /** Verbatim `<picture r:id="…"/>` element (background picture). */
  readonly pictureXml?: string;
  /** Verbatim `<oleObjects>…</oleObjects>` block. */
  readonly oleObjectsXml?: string;
  /** Verbatim `<controls>…</controls>` block (form controls / ActiveX). */
  readonly controlsXml?: string;
}

/**
 * Subset of Excel chart kinds we render natively. Existing chart
 * parts in the source workbook are *not* coerced into this union —
 * they round-trip opaquely and we draw a placeholder instead.
 */
export type ChartKind = "column" | "bar" | "line" | "pie";

/**
 * Named color palettes the renderer cycles through for series.
 *
 * Round-trips by name (not by raw colors) so tweaking a palette's
 * hex values later automatically flows to existing charts. `default`
 * is the historical Office-blue lead palette; `vibrant` /
 * `pastel` / `warm` / `cool` are theme-style alternatives, and
 * `mono` is a single-hue ramp for clean print-style charts.
 */
export type ChartPalette = "default" | "vibrant" | "pastel" | "warm" | "cool" | "mono";

/**
 * Free-floating chart on a worksheet (C15).
 *
 * Same anchor model as {@link SheetImage}: a from-cell + pixel
 * offset and a CSS-pixel size. The renderer reads `dataRange` out
 * of the sheet's cells at draw time so the chart auto-updates
 * when the underlying values change.
 *
 * All `style` fields are optional so charts authored before they
 * existed render with the historical defaults. The renderer treats
 * `undefined` as "use the default" (legend on, gridlines on,
 * data labels off, palette = `"default"`).
 */
export interface SheetChart {
  readonly id: NodeId;
  readonly kind: ChartKind;
  /** A1 range covering header row + body rows of the chart's data. */
  readonly dataRange: string;
  /** When `true`, the first row of `dataRange` is treated as a header. */
  readonly hasHeaderRow: boolean;
  /** When `true`, the first column of `dataRange` provides category labels. */
  readonly hasCategoryColumn: boolean;
  readonly title?: string;
  readonly anchor: import("./drawings.js").ImageAnchor;
  /** Series color theme; defaults to `"default"` when omitted. */
  readonly palette?: ChartPalette;
  /** Hide the legend by setting to `false`; default is shown when ≥2 series. */
  readonly showLegend?: boolean;
  /** When `true`, paint the value next to each bar/point/slice. Default `false`. */
  readonly showDataLabels?: boolean;
  /** When `false`, suppress the value-axis gridlines. Default `true`. */
  readonly showGridlines?: boolean;
  /** Optional value-axis label (e.g. "Revenue"). Renders along the Y for column/line, X for bar. */
  readonly yAxisTitle?: string;
  /** Optional category-axis label (e.g. "Quarter"). Renders along the X for column/line, Y for bar. */
  readonly xAxisTitle?: string;
}

/**
 * Excel Table (ListObject). Mirrors the subset of `xl/tables/*.xml`
 * we read + emit. The OOXML element supports a *lot* more
 * (calculated columns, totals row functions, custom xml mappings…),
 * but those round-trip via {@link TableDef.opaqueXml} — the parser
 * holds the original bytes so we re-emit unchanged tables verbatim.
 *
 * Spec: ECMA-376 §18.5.
 */
export interface TableDef {
  readonly id: NodeId;
  /** OOXML `id` attr — sheet-scoped numeric id used by formulas. */
  readonly tableId: string;
  /** OOXML `name` attr — internal name (matches `displayName` by default). */
  readonly name: string;
  /** OOXML `displayName` attr — what shows in the Name Box / formula refs. */
  readonly displayName: string;
  /** A1 range covered by the table, headers + body + totals row. */
  readonly range: string;
  /** OOXML `headerRowCount` attr. `0` = no headers (rare); `1` = standard. */
  readonly headerRowCount: number;
  /** OOXML `totalsRowCount` attr. `0` = no totals row; `1` = standard. */
  readonly totalsRowCount: number;
  /**
   * Column display names, in `range` column order. We always read at
   * least the names from `<tableColumns>` so column-style banding
   * can label them correctly even when the header cells are blank.
   */
  readonly columnNames: ReadonlyArray<string>;
  /** OOXML style descriptor — kept opaque to round-trip the styling pack. */
  readonly styleInfoXml?: string;
  /**
   * OOXML autoFilter range, if the table has filter buttons enabled.
   * We default to the table's own range when adding a new table; when
   * round-tripping we preserve whatever the source file had.
   */
  readonly autoFilterRange?: string;
  /**
   * Verbatim original `<table …>…</table>` XML. Re-emitted byte-
   * identically when the table is *not* dirty. Cleared on edits so
   * the typed serializer takes over.
   */
  readonly opaqueXml?: string;
  /**
   * OOXML part path (`xl/tables/tableN.xml`). Stable across the
   * session so sheet rels can target it.
   */
  readonly partPath: string;
  /**
   * Relationship id used by the sheet's rels part to reference the
   * table part. Stable across the session so rels round-trip cleanly.
   */
  readonly relId: string;
}

/**
 * Frozen-pane configuration. See `Sheet.freeze` for the OOXML
 * mapping. Both axes default to `0` (no freeze on that axis).
 */
export interface FreezePanes {
  readonly rows: number;
  readonly cols: number;
}

/**
 * Visual override emitted by a conditional-formatting rule when
 * its predicate matches a cell. Sparse — only the fields the rule
 * needs to set are populated; everything else falls back to the
 * cell's base style.
 */
export interface ConditionalFormatOverlay {
  /** Background fill, RRGGBB hex without `#`. */
  readonly fill?: string;
  /** Font colour, RRGGBB hex without `#`. */
  readonly fontColor?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  /** Bar fill for `dataBar` rules, RRGGBB without `#`. */
  readonly barColor?: string;
  /** 0..1 — relative width of the data-bar overlay. */
  readonly barFraction?: number;
}

/**
 * Discriminated union of supported conditional-formatting rule
 * kinds. Mirrors Excel's most-used CF dialog entries; OOXML
 * round-trip is deferred to a future pass.
 */
export type ConditionalFormat =
  | {
      readonly kind: "cellIs";
      readonly id: string;
      /** A1 ranges (multi-area allowed, comma-separated). */
      readonly range: string;
      readonly op: "gt" | "ge" | "lt" | "le" | "eq" | "ne" | "between" | "notBetween";
      /** Single threshold for unary operators; lower bound for between. */
      readonly value: number;
      /** Upper bound for `between` / `notBetween`. */
      readonly value2?: number;
      readonly overlay: ConditionalFormatOverlay;
    }
  | {
      readonly kind: "top10";
      readonly id: string;
      readonly range: string;
      readonly bottom: boolean;
      readonly percent: boolean;
      readonly rank: number;
      readonly overlay: ConditionalFormatOverlay;
    }
  | {
      readonly kind: "containsText";
      readonly id: string;
      readonly range: string;
      readonly text: string;
      /** True for "contains", false for "does not contain". */
      readonly contains: boolean;
      readonly overlay: ConditionalFormatOverlay;
    }
  | {
      readonly kind: "duplicate";
      readonly id: string;
      readonly range: string;
      readonly unique: boolean;
      readonly overlay: ConditionalFormatOverlay;
    }
  | {
      readonly kind: "colorScale";
      readonly id: string;
      readonly range: string;
      /** Min stop colour (RRGGBB). */
      readonly minColor: string;
      /** Optional midpoint colour for 3-stop scales. */
      readonly midColor?: string;
      /** Max stop colour. */
      readonly maxColor: string;
    }
  | {
      readonly kind: "dataBar";
      readonly id: string;
      readonly range: string;
      readonly color: string;
    };

/**
 * Data-validation rule. C11 currently models the `list` kind
 * (in-cell dropdown picker) typed; everything else (whole, decimal,
 * date, textLength, custom) preserves through the opaque round-trip
 * path on `Sheet.opaqueDataValidations` until we extend the model.
 */
export type DataValidation = {
  readonly kind: "list";
  readonly id: string;
  /** A1 range string (single area; multi-area authoring stays as opaque). */
  readonly range: string;
  /**
   * List source. Either a literal comma-separated string of values
   * (`"Yes,No,Maybe"`) or a formula reference (`"=Sheet1!$A$1:$A$5"`).
   * `formula=true` flips the OOXML `formula1` element from a
   * quoted-literal to a reference.
   */
  readonly source: string;
  readonly formula: boolean;
  /** Show the in-cell dropdown arrow (Excel default = true). */
  readonly showDropDown: boolean;
  /** Reject values outside the list (Excel "Stop" style). Default = true. */
  readonly stopOnInvalid: boolean;
  /** Allow the cell to be empty even if the list doesn't include "". */
  readonly allowBlank: boolean;
};

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
  /**
   * Verbatim `<text>…</text>` element captured from the parsed
   * comments part. Carries the full rich-text run tree (`<r><rPr>…
   * </rPr><t>…</t></r>`, multiple runs, embedded fonts, colors).
   * The serializer re-emits this blob when the comment's plain
   * `text` field is byte-identical to the flattened text in the
   * blob — i.e., when the user hasn't actually edited the comment
   * body. When `text` diverges, the serializer falls back to a
   * single-run `<r><t>` for the new text and the rich formatting is
   * (intentionally) dropped.
   */
  readonly textXml?: string;
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
  /**
   * Formula kind:
   *  - `"normal"` (default, omitted): an ordinary per-cell formula.
   *  - `"shared"`: one of N cells participating in a shared formula
   *    group identified by {@link sharedIndex}. Only the master
   *    cell carries `text`; SheetJS expands the followers' text on
   *    parse so all cells in our typed model carry the resolved
   *    formula. The serializer re-emits `<f t="shared" si=… ref=…>`
   *    on the master and `<f t="shared" si=…/>` on followers so the
   *    compact OOXML encoding is preserved.
   *  - `"array"`: cell is part of a CSE-array (legacy `Ctrl+Shift+
   *    Enter`) or modern dynamic-array spill. The master cell holds
   *    the formula body; followers reference the same `ref`.
   */
  readonly kind?: "normal" | "shared" | "array";
  /**
   * Shared-formula group index from the source `<f si="…"/>` attr.
   * Only meaningful when {@link kind} is `"shared"`.
   */
  readonly sharedIndex?: number;
  /**
   * Spill range for shared / array formulas, exactly as written in
   * the source `<f ref="…"/>` attribute (A1 notation). Present on
   * the master cell of the group. Followers are detected by
   * matching {@link sharedIndex}.
   */
  readonly ref?: string;
  /**
   * True when this cell is the master of its shared / array group
   * (the cell that carried the formula body in the source XML).
   * Followers re-emit `<f t="shared" si=…/>` (no body) so Excel
   * resolves them via the master.
   */
  readonly isMaster?: boolean;
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
    pivotTables: false,
  };
}
