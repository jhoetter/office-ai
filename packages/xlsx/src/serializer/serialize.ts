import { ooxml } from "@officeai/core";
import * as XLSX from "@e965/xlsx";
import { EXTENSION_BY_CONTENT_TYPE } from "../model/drawings.js";
import { formatA1 } from "../model/refs.js";
import type {
  AutoFilter,
  Cell,
  ConditionalFormat,
  CustomFilterOp,
  DataValidation,
  FilterColumn,
  FreezePanes,
  Sheet,
  XlsxSnapshot,
  XlsxWorkbook,
} from "../model/types.js";
import { CHART_CONTENT_TYPE, mintChartPartPath, serializeChartPart } from "./charts.js";
import { serializeCommentsPart } from "./comments.js";
import {
  buildDrawingRels,
  DRAWING_CONTENT_TYPE,
  injectDrawingRef,
  mintDrawingPartPath,
  serializeDrawingPart,
  upsertSheetDrawingRel,
} from "./drawings.js";
import { XlsxSerializeError } from "./errors.js";
import { serializePivotParts } from "./pivot-tables.js";
import { syncSheetToSheetJS } from "./sheet-sync.js";
import { serializeStylesXml } from "./styles.js";

const WORKBOOK_PART = "xl/workbook.xml";
const STYLES_PART = "xl/styles.xml";
const WORKSHEET_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const WORKSHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";

/**
 * Serialize an `XlsxSnapshot` back to bytes.
 *
 * Phase 5+ contract:
 *   - No dirty flags set → byte-content-identical re-emit. Untouched
 *     workbooks always round-trip exactly.
 *   - Dirty sheets → for each dirty sheet, sync the typed cells +
 *     merges back onto the SheetJS WorkSheet, then ask SheetJS to
 *     emit a single-sheet workbook for that sheet, and substitute the
 *     emitted `xl/worksheets/sheetN.xml` into the master container.
 *     Brand-new sheets reach this same path because the handler has
 *     already added them to `book.SheetNames` + `book.Sheets`.
 *   - Dirty workbook / rels / contentTypes → re-emit the affected
 *     parts from the typed `XlsxWorkbook.sheets` array. Renames,
 *     insertions, and the matching workbook-rels + content-types
 *     overrides all flow through these three rewrites.
 *
 * Trade-offs documented in `docs/build-log/xlsx.md`:
 *   - String cells written by Phase 5 commands are emitted inline
 *     (`t="inlineStr"`) so the shared-strings part is never disturbed.
 *   - Phase 5 does not author formulas; passing a formula through the
 *     model preserves the formula text and any cached value, but does
 *     not recompute. Phase 7 adds the formula engine + recalc.
 *
 * Phase 6+ migrates to a native sheet-XML emitter that surgically
 * patches `<sheetData>` + `<mergeCells>` while preserving all other
 * worksheet XML (cols, sheetView, conditional formatting, etc.). For
 * Phase 5 the SheetJS-based path is sufficient for our fixtures.
 */
export async function serializeXlsx(snapshot: XlsxSnapshot): Promise<ArrayBuffer> {
  const container = snapshot.container.clone();

  const dirty = snapshot.dirty;
  const dirtySheetPaths = dirty.sheets;
  const unsupportedDirty = dirty.sharedStrings || dirty.threadedComments.size > 0;

  if (unsupportedDirty) {
    throw new XlsxSerializeError(
      "container-failed",
      "Serializer supports `sheets` + `workbook` + `styles` + `rels` + `contentTypes` + `comments` + `sheetRels`; sst/threadedComments rewrites land in later phases"
    );
  }

  // Drawings + media are processed BEFORE the sheet pass so the sheet
  // pass can splice the freshly-minted `<drawing r:id>` reference
  // straight into the regenerated worksheet XML.
  const drawingRidByPath = new Map<string, string | null>();
  const emittedChartParts = new Set<string>();
  if (dirty.drawings.size > 0) {
    rewriteDirtyDrawings(snapshot.root, container, dirty.drawings, drawingRidByPath, emittedChartParts);
    dropOrphanChartParts(container, emittedChartParts);
  }
  if (dirty.media.size > 0) {
    rewriteDirtyMedia(snapshot.root, container, dirty.media);
  }
  if (dirty.removedMediaParts.size > 0) {
    dropRemovedMediaParts(container, dirty.removedMediaParts);
  }

  if (dirtySheetPaths.size > 0) {
    await rewriteDirtySheets(snapshot, container, dirtySheetPaths, drawingRidByPath);
  }

  if (dirty.comments.size > 0) {
    rewriteDirtyComments(snapshot.root, container, dirty.comments);
  }

  if (dirty.sheetRels.size > 0) {
    rewriteDirtySheetRels(snapshot.root, container, dirty.sheetRels);
  }

  if (dirty.removedSheetParts.size > 0) {
    dropRemovedSheetParts(container, dirty.removedSheetParts);
  }

  if (dirty.contentTypes) {
    rewriteContentTypes(snapshot.root, container, dirty.removedSheetParts, dirty.removedMediaParts);
  }

  if (dirty.rels) {
    rewriteWorkbookRels(snapshot.root, container, dirty.removedSheetParts);
  }

  if (dirty.workbook) {
    rewriteWorkbookSheets(snapshot.root, container);
  }

  if (dirty.styles) {
    rewriteStylesXml(snapshot.root, container);
  }

  // F1 Phase 1 — pivot tables and caches re-emit byte-identical
  // from `raw`. Phase 3 will gate this on `dirty.pivotTables` for
  // typed re-renders; today the call is unconditional and idempotent
  // (the bytes are equal to what's already in the cloned container).
  serializePivotParts(snapshot.root, container);

  try {
    return await container.serialize();
  } catch (err) {
    throw new XlsxSerializeError("container-failed", "Failed to re-emit OOXML container", { cause: err });
  }
}

async function rewriteDirtySheets(
  snapshot: XlsxSnapshot,
  master: ooxml.OoxmlContainer,
  dirtySheetPaths: ReadonlySet<string>,
  drawingRidByPath: ReadonlyMap<string, string | null>
): Promise<void> {
  const book = snapshot.root.sheetjs;
  const sheetsByPath = new Map<string, Sheet>();
  for (const sheet of snapshot.root.sheets) sheetsByPath.set(sheet.partPath, sheet);

  for (const path of dirtySheetPaths) {
    const sheet = sheetsByPath.get(path);
    if (!sheet) {
      throw new XlsxSerializeError("sheet-failed", `dirty sheet path not found in workbook: ${path}`, {
        partPath: path,
      });
    }
    const ws = book.Sheets[sheet.name];
    if (!ws) {
      throw new XlsxSerializeError("sheet-failed", `SheetJS missing worksheet "${sheet.name}"`, {
        partPath: path,
      });
    }
    syncSheetToSheetJS(sheet, ws);
  }

  let emitted: ArrayBuffer;
  try {
    emitted = XLSX.write(book, {
      type: "array",
      bookType: "xlsx",
      bookSST: false,
      compression: true,
      cellDates: false,
    }) as ArrayBuffer;
  } catch (err) {
    throw new XlsxSerializeError("sheet-failed", "SheetJS write failed", { cause: err });
  }

  let emittedContainer: ooxml.OoxmlContainer;
  try {
    emittedContainer = await ooxml.OoxmlContainer.load(emitted);
  } catch (err) {
    throw new XlsxSerializeError("sheet-failed", "Failed to read SheetJS-emitted xlsx", { cause: err });
  }

  const orderedSheets = snapshot.root.sheets.filter((s) => s.kind === "worksheet");
  for (const path of dirtySheetPaths) {
    const sheet = sheetsByPath.get(path);
    if (!sheet) continue;
    const sheetIdx = orderedSheets.findIndex((s) => s.partPath === path);
    if (sheetIdx === -1) {
      throw new XlsxSerializeError("sheet-failed", `dirty sheet not found among worksheet sheets: ${path}`, {
        partPath: path,
      });
    }
    const emittedPath = `xl/worksheets/sheet${sheetIdx + 1}.xml`;
    if (!emittedContainer.has(emittedPath)) {
      throw new XlsxSerializeError(
        "sheet-failed",
        `SheetJS output missing expected sheet part ${emittedPath}`,
        { partPath: emittedPath }
      );
    }
    let xml = emittedContainer.readText(emittedPath);
    xml = injectStyleIds(xml, sheet.cells);
    xml = injectFormulas(xml, sheet.cells);
    xml = injectHiddenRows(xml, sheet.hiddenRows);
    xml = injectAutoFilter(xml, sheet.autoFilter);
    xml = injectSheetViews(xml, sheet);
    xml = injectTabColor(xml, sheet);
    xml = injectCols(xml, sheet);
    xml = injectConditionalFormats(xml, sheet);
    xml = injectDataValidations(xml, sheet.dataValidations, sheet.opaqueDataValidations);
    xml = injectHyperlinks(xml, sheet.hyperlinksXml);
    xml = injectTableParts(xml, sheet);
    xml = injectOpaqueTail(xml, sheet);
    if (drawingRidByPath.has(path)) {
      xml = injectDrawingRef(xml, drawingRidByPath.get(path) ?? null);
    } else if (sheet.drawingPartPath) {
      // Sheet has a pre-existing drawing we did not re-author; reuse
      // the existing rId from the on-disk sheet rels so the
      // SheetJS-emitted XML keeps pointing at it.
      const rels = ooxml.RelationshipGraph.loadFor(master, path);
      const existing = rels.relationships.find(
        (r) => r.type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
      );
      xml = injectDrawingRef(xml, existing?.id ?? null);
    }
    master.writeText(path, xml);
  }
}

/**
 * For each sheet whose `dirty.drawings` flag is set, (re)emit its
 * `xl/drawings/drawingN.xml`, that part's rels, and the sheet rels
 * that point at it. Tracks the rId Excel will see in the worksheet's
 * `<drawing>` element so the sheet pass can splice it in.
 */
function rewriteDirtyDrawings(
  workbook: XlsxWorkbook,
  container: ooxml.OoxmlContainer,
  paths: ReadonlySet<string>,
  drawingRidByPath: Map<string, string | null>,
  emittedChartParts: Set<string>
): void {
  for (const sheetPartPath of paths) {
    const sheet = workbook.sheets.find((s) => s.partPath === sheetPartPath);
    if (!sheet) continue;

    const sheetRels = ooxml.RelationshipGraph.loadFor(container, sheetPartPath);
    const hasDrawables = sheet.images.length > 0 || sheet.charts.length > 0;

    if (!hasDrawables) {
      // Nothing left → drop the drawing part + rels and clear the
      // sheet's drawing relationship. Chart parts that the dropped
      // drawing was pointing at are reaped lazily — the next save
      // sees them as orphaned in the container and they get cleaned
      // up the same way orphan media is, in `dropOrphanChartParts`.
      if (sheet.drawingPartPath) {
        if (container.has(sheet.drawingPartPath)) container.removePart(sheet.drawingPartPath);
        const drawingRelsPath = ooxml.RelationshipGraph.relsPathFor(sheet.drawingPartPath);
        if (container.has(drawingRelsPath)) container.removePart(drawingRelsPath);
      }
      upsertSheetDrawingRel(sheetRels, sheetPartPath, null);
      sheetRels.writeBack(container);
      drawingRidByPath.set(sheetPartPath, null);
      continue;
    }

    const drawingPartPath = sheet.drawingPartPath ?? mintDrawingPartPath(container);

    // Mint a fresh chart-part path per typed chart. We don't try to
    // reuse the chart's old slot across saves because the typed
    // model doesn't carry a `partPath` (deliberately — chart parts
    // are derived state). The previously-written chart part files
    // in this slot get overwritten if the path collides; otherwise
    // `dropOrphanChartParts` reaps the stale ones.
    const chartDescs: Array<{ readonly id: string; readonly chartPartPath: string }> = [];
    for (const chart of sheet.charts) {
      const chartPartPath = mintChartPartPath(container, emittedChartParts);
      emittedChartParts.add(chartPartPath);
      const chartXml = serializeChartPart(chart, sheet.name);
      container.writeText(chartPartPath, chartXml);
      // Charts authored by us do not (yet) reference embedded data
      // sources, themes, or images — they only need a presence in
      // the drawing rels graph below. The chart-rels file is
      // therefore omitted entirely; Office tolerates a missing
      // `xl/charts/_rels/chartN.xml.rels` for stand-alone charts.
      chartDescs.push({ id: chart.id, chartPartPath });
    }

    const {
      graph: drawingRels,
      embedRidByMediaPath,
      chartRidByChartId,
    } = buildDrawingRels(drawingPartPath, sheet.images, workbook.images, chartDescs);
    const xml = serializeDrawingPart(sheet.images, embedRidByMediaPath, sheet.charts, chartRidByChartId);
    container.writeText(drawingPartPath, xml);
    drawingRels.writeBack(container);

    const rid = upsertSheetDrawingRel(sheetRels, sheetPartPath, drawingPartPath);
    sheetRels.writeBack(container);
    drawingRidByPath.set(sheetPartPath, rid);
  }
}

/**
 * After every drawing has been (re)written, prune any `xl/charts/*.xml`
 * part that's no longer referenced from any sheet's drawing rels.
 * This keeps brand-new charts from accreting on each save when a
 * previous save authored a chart that has since been removed via
 * `xlsx:remove-chart`.
 */
function dropOrphanChartParts(container: ooxml.OoxmlContainer, emittedChartParts: ReadonlySet<string>): void {
  const live = new Set<string>(emittedChartParts);
  // Anything we did NOT emit this round AND that lives under
  // xl/charts/ AND is not referenced from any drawing's rels
  // graph is by definition orphaned.
  const allParts: string[] = [];
  for (const path of container.parts.keys()) {
    if (path.startsWith("xl/charts/") && /chart\d+\.xml$/.test(path)) {
      allParts.push(path);
    }
  }
  if (allParts.length === 0) return;

  // Reference scan: walk every drawing-rels file and collect targets.
  for (const path of container.parts.keys()) {
    if (!path.startsWith("xl/drawings/_rels/") || !path.endsWith(".xml.rels")) continue;
    const text = container.readText(path);
    // We rely on the relationship-graph parser indirectly via a
    // very small regex here — the only thing that matters is that
    // a chart's filename appears as a Target. Re-loading every rel
    // graph just to read targets would be a lot of allocation for
    // a cleanup pass.
    for (const m of text.matchAll(/Target="([^"]+)"/g)) {
      const t = m[1] ?? "";
      // Resolve relative to `xl/drawings/` (the rels owner's dir).
      const ownerDir = "xl/drawings";
      const resolved = resolveRelTarget(ownerDir, t);
      if (resolved && resolved.startsWith("xl/charts/")) live.add(resolved);
    }
  }

  for (const part of allParts) {
    if (!live.has(part)) {
      container.removePart(part);
      const relsPath = ooxml.RelationshipGraph.relsPathFor(part);
      if (container.has(relsPath)) container.removePart(relsPath);
    }
  }
}

function resolveRelTarget(ownerDir: string, target: string): string | null {
  if (!target) return null;
  if (target.startsWith("/")) return target.slice(1);
  const segs = ownerDir.split("/").filter(Boolean);
  for (const part of target.split("/")) {
    if (part === "..") segs.pop();
    else if (part === "." || part === "") continue;
    else segs.push(part);
  }
  return segs.join("/");
}

/**
 * (Re)write media bytes for every dirty media path.
 */
function rewriteDirtyMedia(
  workbook: XlsxWorkbook,
  container: ooxml.OoxmlContainer,
  paths: ReadonlySet<string>
): void {
  for (const path of paths) {
    const blob = workbook.images.get(path);
    if (!blob) continue;
    container.writeBytes(path, blob.bytes);
  }
}

function dropRemovedMediaParts(container: ooxml.OoxmlContainer, paths: ReadonlySet<string>): void {
  for (const path of paths) {
    if (container.has(path)) container.removePart(path);
  }
}

/**
 * Set / strip `hidden="1"` on `<row>` elements based on
 * `sheet.hiddenRows`. Rows present in the set get `hidden="1"`,
 * `ht="0"`, and `customHeight="1"` (matches Excel's filter-driven
 * row hiding wire format). Rows not in the set get those three
 * attributes stripped so re-applying then clearing a filter
 * round-trips cleanly.
 *
 * Rows that aren't already emitted as `<row>` elements but live in
 * the hidden set (rare — the row would have to be empty) are not
 * synthesised here; the autoFilter evaluator only adds rows that
 * carry data, and SheetJS emits a `<row>` for every cell it sees.
 */
function injectHiddenRows(xml: string, hiddenRows: ReadonlySet<number>): string {
  if (xml.indexOf("<row") === -1) return xml;
  return xml.replace(/<row\b([^/>]*?)(\/?)>/g, (_match, attrs: string, selfClose: string) => {
    const refMatch = /\br=("|')([^"']+)\1/.exec(attrs);
    if (!refMatch) return `<row${attrs}${selfClose}>`;
    const rowNum = Number(refMatch[2]);
    if (!Number.isInteger(rowNum) || rowNum < 1) return `<row${attrs}${selfClose}>`;
    const rowIdx = rowNum - 1;
    let stripped = attrs
      .replace(/\s+hidden=("|')[^"']*\1/, "")
      .replace(/\s+ht=("|')[^"']*\1/, "")
      .replace(/\s+customHeight=("|')[^"']*\1/, "");
    if (hiddenRows.has(rowIdx)) {
      stripped = `${stripped} hidden="1" ht="0" customHeight="1"`;
    }
    return `<row${stripped}${selfClose}>`;
  });
}

/**
 * Strip any pre-existing `<autoFilter>` block from the worksheet XML
 * and, if `autoFilter` is non-undefined, splice a freshly serialized
 * one in immediately before `</worksheet>` (Excel orders it after
 * `<sheetData>` / `<mergeCells>`).
 */
function injectAutoFilter(xml: string, autoFilter: AutoFilter | undefined): string {
  const next = xml.replace(/<autoFilter\b[^>]*(?:\/>|>[\s\S]*?<\/autoFilter>)/g, "");
  if (!autoFilter) return next;

  const ref = `${colToA1(autoFilter.range.c1)}${autoFilter.range.r1 + 1}:${colToA1(autoFilter.range.c2)}${autoFilter.range.r2 + 1}`;
  const cols: string[] = [];
  // Sort by colId for deterministic output.
  const ids = [...autoFilter.columns.keys()].sort((a, b) => a - b);
  for (const id of ids) {
    const fc = autoFilter.columns.get(id);
    if (!fc) continue;
    cols.push(serializeFilterColumn(id, fc));
  }
  const block = `<autoFilter ref="${ref}">${cols.join("")}</autoFilter>`;

  const closeIdx = next.lastIndexOf("</worksheet>");
  if (closeIdx === -1) {
    // Defensive: append the block at the end. SheetJS always emits a
    // `<worksheet>` wrapper so this branch is essentially unreachable.
    return next + block;
  }
  return next.slice(0, closeIdx) + block + next.slice(closeIdx);
}

/**
 * Strip any pre-existing `<sheetViews>` block, then re-emit one that
 * carries our `<pane state="frozen"/>` (when the sheet has a freeze)
 * or no `<pane>` at all (when it doesn't).
 *
 * We rebuild the `<sheetViews>` because SheetJS may emit one with a
 * stale `<pane>` element from the original file, and surgically
 * patching the existing block (preserving its other attrs) gets
 * gnarly fast — `<sheetView>` carries optional `tabSelected`,
 * `zoomScale`, `view`, and a `<selection>` child that we'd otherwise
 * have to round-trip individually. The simpler "drop and re-emit"
 * approach is acceptable because users overwhelmingly stick to
 * default view settings; freeze panes is the only one we surface
 * via the toolbar.
 *
 * Excel's `topLeftCell` is the cell that sits in the top-left of the
 * scrolling pane (i.e. one row below `ySplit` and one column right of
 * `xSplit`). We compute it deterministically rather than preserve
 * whatever Excel last wrote, so the output is stable.
 */
/**
 * C10 — Re-inject opaque `<conditionalFormatting>` blocks captured
 * by the parser. The parser stores them verbatim per-sheet so we
 * can preserve every CF rule (data bars, icon sets, color scales,
 * formula rules, etc.) on dirty round-trip without modelling them
 * fully. Typed authoring lives in `Sheet.conditionalFormats` and
 * is not yet emitted (deferred to a future pass).
 */
function injectConditionalFormats(xml: string, sheet: Sheet): string {
  // First, drop any pre-existing CF blocks so we don't double-emit
  // when SheetJS already echoed them through.
  const stripped = xml.replace(
    /<conditionalFormatting\b[^>]*?(?:\/>|>[\s\S]*?<\/conditionalFormatting>)/g,
    ""
  );
  // Re-emit captured opaque blocks (so existing rules from the
  // source file survive a dirty save) AND synthesised blocks for
  // typed `conditionalFormats` so newly authored rules now land in
  // the OOXML output. Typed rules without an overlay (color scale,
  // data bar) carry their full styling inline; rules WITH an
  // overlay (cellIs, top10, containsText, duplicate) reference
  // `dxfId="0"` — Excel renders them as the rule with no styling
  // when dxfId 0 isn't a populated dxf entry, which preserves the
  // rule semantics even though the colour overlay is dropped. A
  // later pass will allocate dxf entries in `xl/styles.xml` to
  // restore the overlay paint; the spec for that work is the dxf
  // emission TODO in the office-roundtrip-gaps audit plan.
  const opaqueXml = sheet.opaqueConditionalFormats.join("");
  const typedXml = sheet.conditionalFormats.map(renderTypedConditionalFormat).join("");
  const block = opaqueXml + typedXml;
  if (block.length === 0) return stripped;
  const closeIdx = stripped.lastIndexOf("</worksheet>");
  if (closeIdx === -1) return stripped + block;
  return stripped.slice(0, closeIdx) + block + stripped.slice(closeIdx);
}

function renderTypedConditionalFormat(rule: ConditionalFormat): string {
  const sqref = escapeXmlAttr(rule.range);
  // Priority is set to 1 for every typed rule. Real Excel stacks
  // rules by user-authored priority; we don't currently track
  // authoring order across the typed and opaque sets, so collisions
  // resolve by source order ("first cfRule wins"). This is
  // acceptable for the rule kinds we model (no overlapping
  // priorities cause silent drops).
  switch (rule.kind) {
    case "cellIs":
      return (
        `<conditionalFormatting sqref="${sqref}">` +
        `<cfRule type="cellIs" dxfId="0" priority="1" operator="${cellIsOpToOoxml(rule.op)}">` +
        `<formula>${escapeXmlText(String(rule.value))}</formula>` +
        (rule.value2 !== undefined ? `<formula>${escapeXmlText(String(rule.value2))}</formula>` : "") +
        `</cfRule></conditionalFormatting>`
      );
    case "top10": {
      const bottom = rule.bottom ? ' bottom="1"' : "";
      const percent = rule.percent ? ' percent="1"' : "";
      return (
        `<conditionalFormatting sqref="${sqref}">` +
        `<cfRule type="top10" dxfId="0" priority="1"${bottom}${percent} rank="${rule.rank}"/>` +
        `</conditionalFormatting>`
      );
    }
    case "containsText": {
      const op = rule.contains ? "containsText" : "notContainsText";
      const fn = rule.contains ? "ISNUMBER(SEARCH" : "NOT(ISNUMBER(SEARCH";
      const closer = rule.contains ? "))" : ")))";
      // First sqref segment is the anchor cell for the formula.
      const anchor = anchorOfSqref(rule.range);
      const formula = `${fn}("${escapeXmlText(rule.text)}",${anchor})${closer}`;
      return (
        `<conditionalFormatting sqref="${sqref}">` +
        `<cfRule type="${op}" dxfId="0" priority="1" operator="${op}" text="${escapeXmlAttr(rule.text)}">` +
        `<formula>${formula}</formula>` +
        `</cfRule></conditionalFormatting>`
      );
    }
    case "duplicate": {
      const type = rule.unique ? "uniqueValues" : "duplicateValues";
      return (
        `<conditionalFormatting sqref="${sqref}">` +
        `<cfRule type="${type}" dxfId="0" priority="1"/>` +
        `</conditionalFormatting>`
      );
    }
    case "colorScale": {
      const stops = rule.midColor
        ? `<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>`
        : `<cfvo type="min"/><cfvo type="max"/>`;
      const colors = rule.midColor
        ? `<color rgb="FF${rule.minColor}"/><color rgb="FF${rule.midColor}"/><color rgb="FF${rule.maxColor}"/>`
        : `<color rgb="FF${rule.minColor}"/><color rgb="FF${rule.maxColor}"/>`;
      return (
        `<conditionalFormatting sqref="${sqref}">` +
        `<cfRule type="colorScale" priority="1">` +
        `<colorScale>${stops}${colors}</colorScale>` +
        `</cfRule></conditionalFormatting>`
      );
    }
    case "dataBar":
      return (
        `<conditionalFormatting sqref="${sqref}">` +
        `<cfRule type="dataBar" priority="1">` +
        `<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF${rule.color}"/></dataBar>` +
        `</cfRule></conditionalFormatting>`
      );
    default: {
      const _exhaustive: never = rule;
      void _exhaustive;
      return "";
    }
  }
}

/**
 * Map our short typed `cellIs` operator codes to the OOXML
 * `ST_ConditionalFormattingOperator` token names. Excel writes
 * the long form (`greaterThan`, `lessThanOrEqual`, …) so we
 * follow suit; otherwise opening the file would surface a "rule
 * not understood" warning.
 */
function cellIsOpToOoxml(op: "gt" | "ge" | "lt" | "le" | "eq" | "ne" | "between" | "notBetween"): string {
  switch (op) {
    case "gt":
      return "greaterThan";
    case "ge":
      return "greaterThanOrEqual";
    case "lt":
      return "lessThan";
    case "le":
      return "lessThanOrEqual";
    case "eq":
      return "equal";
    case "ne":
      return "notEqual";
    case "between":
      return "between";
    case "notBetween":
      return "notBetween";
  }
}

function anchorOfSqref(sqref: string): string {
  // Take the first cell of the first range. `sqref` may carry
  // multiple space-separated ranges (e.g. "A1:A5 C1:C5"); we anchor
  // on the top-left of the leftmost range so the conditional
  // formula evaluates relative to a real cell. Stripping `$`
  // anchors keeps the output identical to what Excel itself emits
  // for these CF formulas.
  const first = sqref.split(/\s+/)[0] ?? sqref;
  const cell = first.split(":")[0] ?? first;
  return cell.replace(/\$/g, "");
}

/**
 * C11 — Re-emit `<dataValidations>` for dirty sheets.
 *
 * Strategy mirrors the conditional-format path:
 *   1. Drop any pre-existing `<dataValidations>` block from the
 *      SheetJS output (it ignores them, but we want a clean slate).
 *   2. If the user has typed `list` rules OR the parser captured
 *      non-list rules opaquely, render a fresh `<dataValidations>`
 *      containing both the typed list entries and the opaque
 *      `<dataValidation>` children verbatim.
 *
 * Excel mandates the `<dataValidations>` block sit between
 * `<mergeCells>` (or `<phoneticPr>`) and `<hyperlinks>`. The lazy
 * placement we use — right before `</worksheet>` — is also accepted
 * by Excel and matches what `injectConditionalFormats` does.
 */
function injectDataValidations(
  xml: string,
  typed: ReadonlyArray<DataValidation>,
  opaque: string | undefined
): string {
  const stripped = xml.replace(/<dataValidations\b[^>]*>[\s\S]*?<\/dataValidations>/g, "");
  const typedList = typed.filter((dv) => dv.kind === "list");
  if (typedList.length === 0 && !opaque) return stripped;

  const typedXml = typedList.map(renderListValidation).join("");
  // The opaque block is already a complete `<dataValidations>…</dataValidations>`
  // wrapper. Pull just its inner children so we can re-wrap with a
  // single block that holds both typed + opaque rules.
  const opaqueInner = opaque
    ? opaque.replace(/^<dataValidations\b[^>]*>/, "").replace(/<\/dataValidations>$/, "")
    : "";
  const total = typedList.length + (opaqueInner.match(/<dataValidation\b/g)?.length ?? 0);
  const block = `<dataValidations count="${total}">${typedXml}${opaqueInner}</dataValidations>`;

  const closeIdx = stripped.lastIndexOf("</worksheet>");
  if (closeIdx === -1) return stripped + block;
  return stripped.slice(0, closeIdx) + block + stripped.slice(closeIdx);
}

function renderListValidation(dv: DataValidation): string {
  const sqref = escapeXmlAttr(dv.range);
  const allowBlank = dv.allowBlank ? ' allowBlank="1"' : "";
  const showDropDown = dv.showDropDown ? "" : ' showDropDown="1"';
  const errorStyle = dv.stopOnInvalid ? "" : ' errorStyle="warning"';
  const formula1 = dv.formula
    ? `<formula1>${escapeXmlText(dv.source)}</formula1>`
    : `<formula1>"${escapeXmlText(dv.source)}"</formula1>`;
  return (
    `<dataValidation type="list"${allowBlank}${showDropDown}${errorStyle} sqref="${sqref}">` +
    formula1 +
    `</dataValidation>`
  );
}

/**
 * Non-destructively re-emit `<sheetViews>`.
 *
 * Strategy:
 *   - If the sheet has no captured `sheetViewsXml`, fall back to the
 *     synthetic block (legacy behaviour).
 *   - Otherwise, drop SheetJS's regenerated `<sheetViews>` and splice
 *     the original block back in, then surgically replace just its
 *     `<pane>` element with one derived from `sheet.freeze`. This
 *     preserves zoom, selection, view mode, gridline toggles, and
 *     any other `<sheetView>` children we don't model.
 */
function injectSheetViews(xml: string, sheet: Sheet): string {
  const next = xml.replace(/<sheetViews\b[^>]*>[\s\S]*?<\/sheetViews>/g, "");
  const block = mergeSheetViews(sheet);
  if (!block) return next;

  // Excel orders `<sheetViews>` immediately after `<dimension>` (or
  // first thing in `<worksheet>` if `<dimension>` isn't present).
  const dimMatch = /<dimension\b[^/>]*\/?>(?:<\/dimension>)?/.exec(next);
  if (dimMatch) {
    const insertAt = dimMatch.index + dimMatch[0].length;
    return next.slice(0, insertAt) + block + next.slice(insertAt);
  }
  const wsMatch = /<worksheet\b[^>]*>/.exec(next);
  if (wsMatch) {
    const insertAt = wsMatch.index + wsMatch[0].length;
    return next.slice(0, insertAt) + block + next.slice(insertAt);
  }
  return next + block;
}

function mergeSheetViews(sheet: Sheet): string {
  const original = sheet.sheetViewsXml;
  const freeze = sheet.freeze;

  if (original) {
    return rewritePaneInSheetViews(original, freeze);
  }

  if (!freeze || (freeze.rows <= 0 && freeze.cols <= 0)) {
    return "";
  }

  return (
    `<sheetViews>` +
    `<sheetView tabSelected="1" workbookViewId="0">` +
    renderPane(freeze) +
    `</sheetView>` +
    `</sheetViews>`
  );
}

/**
 * Replace just the `<pane>` element inside an existing `<sheetViews>`
 * block — keep all other `<sheetView>` children verbatim. When the
 * typed freeze field is `undefined` (or zero on both axes) the pane
 * element is stripped entirely; the rest of the original block is
 * preserved.
 */
function rewritePaneInSheetViews(original: string, freeze: FreezePanes | undefined): string {
  const stripped = original.replace(/<pane\b[^/>]*\/?>/g, "");
  if (!freeze || (freeze.rows <= 0 && freeze.cols <= 0)) {
    return stripped;
  }
  const newPane = renderPane(freeze);
  // Splice the new <pane> as the first child of the first <sheetView>.
  const svOpenRe = /<sheetView\b[^>]*>/;
  const svMatch = svOpenRe.exec(stripped);
  if (svMatch) {
    const insertAt = svMatch.index + svMatch[0].length;
    return stripped.slice(0, insertAt) + newPane + stripped.slice(insertAt);
  }
  // Original block had only self-closing `<sheetView/>` elements; rewrite
  // the first one into open-form so we have somewhere to put `<pane>`.
  const selfRe = /<sheetView\b([^/>]*)\/>/;
  const sm = selfRe.exec(stripped);
  if (sm) {
    return (
      stripped.slice(0, sm.index) +
      `<sheetView${sm[1]}>${newPane}</sheetView>` +
      stripped.slice(sm.index + sm[0].length)
    );
  }
  // No `<sheetView>` at all — defensively wrap one.
  const innerOpen = stripped.indexOf(">");
  if (innerOpen === -1) return stripped;
  return (
    stripped.slice(0, innerOpen + 1) +
    `<sheetView tabSelected="1" workbookViewId="0">${newPane}</sheetView>` +
    stripped.slice(innerOpen + 1)
  );
}

function renderPane(freeze: FreezePanes): string {
  const xSplit = Math.max(0, Math.floor(freeze.cols));
  const ySplit = Math.max(0, Math.floor(freeze.rows));
  const topLeftRow = ySplit + 1;
  const topLeftCol = colToA1(xSplit);
  const topLeft = `${topLeftCol}${topLeftRow}`;
  const activePane = xSplit > 0 && ySplit > 0 ? "bottomRight" : xSplit > 0 ? "topRight" : "bottomLeft";
  const xAttr = xSplit > 0 ? ` xSplit="${xSplit}"` : "";
  const yAttr = ySplit > 0 ? ` ySplit="${ySplit}"` : "";
  return `<pane${xAttr}${yAttr} topLeftCell="${topLeft}" activePane="${activePane}" state="frozen"/>`;
}

/**
 * Re-inject the original `<cols>` band on dirty save. SheetJS's emitter
 * regenerates `<cols>` from its own `!cols` array which we do not
 * populate from typed `Sheet.columnWidths`, so the source band would
 * otherwise be dropped. We strip whatever SheetJS produced and splice
 * our captured block back in immediately after `</sheetViews>` (Excel's
 * canonical position for `<cols>`).
 */
function injectCols(xml: string, sheet: Sheet): string {
  const next = xml.replace(/<cols\b[^>]*>[\s\S]*?<\/cols>/g, "");
  const block = sheet.colsXml;
  if (!block) return next;
  return spliceBeforeSheetData(next, block);
}

/**
 * Re-inject the original `<hyperlinks>` block on dirty save. SheetJS
 * does not emit sheet-level hyperlinks, and the matching `r:id` rels
 * survive in the sheet rels graph independently — re-injecting just
 * the body is enough to restore Excel's hyperlink display.
 */
function injectHyperlinks(xml: string, hyperlinksXml: string | undefined): string {
  const next = xml.replace(/<hyperlinks\b[^>]*>[\s\S]*?<\/hyperlinks>/g, "");
  if (!hyperlinksXml) return next;
  return spliceBeforePageMargins(next, hyperlinksXml);
}

/**
 * Re-emit `<tableParts>` from the typed `tables` array on dirty save.
 * Excel requires `count` and inner `<tablePart r:id="…"/>` entries
 * pointing at each `xl/tables/tableN.xml`. We rely on the rels
 * round-trip path to keep the matching rIds alive in the sheet rels.
 *
 * If typed `tables` is empty but we have a captured `tablePartsXml`,
 * we re-emit that opaquely — defensive for files that hit this path
 * before the typed table parser was hardened.
 */
function injectTableParts(xml: string, sheet: Sheet): string {
  const next = xml.replace(/<tableParts\b[^>]*(?:\/>|>[\s\S]*?<\/tableParts>)/g, "");
  let block = "";
  if (sheet.tables.length > 0) {
    const parts = sheet.tables.map((t) => `<tablePart r:id="${escapeXmlAttr(t.relId)}"/>`).join("");
    block = `<tableParts count="${sheet.tables.length}">${parts}</tableParts>`;
  } else if (sheet.tablePartsXml) {
    block = sheet.tablePartsXml;
  }
  if (!block) return next;
  return appendBeforeWorksheetClose(next, block);
}

/**
 * Re-inject opaque page-setup-style children — sheetProtection,
 * pageMargins, pageSetup, printOptions, headerFooter, rowBreaks,
 * colBreaks — that the parser captured verbatim. These are dropped
 * by SheetJS's worksheet emitter today; re-emitting them keeps
 * print configuration and protection state alive across saves.
 *
 * Excel's canonical ordering is:
 *   sheetProtection, autoFilter, sortState, dataConsolidate,
 *   customSheetViews, mergeCells, phoneticPr, conditionalFormatting,
 *   dataValidations, hyperlinks, printOptions, pageMargins,
 *   pageSetup, headerFooter, rowBreaks, colBreaks, customProperties,
 *   cellWatches, ignoredErrors, smartTags, drawing, legacyDrawing,
 *   legacyDrawingHF, picture, oleObjects, controls, webPublishItems,
 *   tableParts, extLst.
 *
 * We splice each block in just before `</worksheet>` in the order
 * above so Excel parses the document without complaint. Some
 * orderings are technically required (sheetProtection must precede
 * autoFilter, etc.) — for the dirty-rewrite path we already wrote
 * autoFilter / mergeCells / etc., so we only inject what comes
 * AFTER those in canonical order, which means everything in this
 * function lands at the tail of the worksheet.
 */
/**
 * Stamp the per-sheet "Tab Color" into the worksheet's `<sheetPr>`.
 *
 * `<sheetPr>` is the very first child of `<worksheet>` per ECMA-376.
 * If SheetJS emitted one already (rare but possible) we splice the
 * `<tabColor>` child into it, replacing any prior tabColor. If not
 * we insert a fresh `<sheetPr>` immediately after the opening
 * `<worksheet ...>` tag. When `sheet.tabColor` is `undefined` we
 * remove any existing `<tabColor>` child but keep `<sheetPr>` (Excel
 * is happy with empty `<sheetPr/>`).
 */
function injectTabColor(xml: string, sheet: Sheet): string {
  const tabColorChild = sheet.tabColor
    ? `<tabColor rgb="${escapeXmlAttr(sheet.tabColor.toUpperCase())}"/>`
    : "";
  const sheetPrMatch = /<sheetPr\b([^>]*)(\/?)>([\s\S]*?)(<\/sheetPr>)?/.exec(xml);
  if (sheetPrMatch && sheetPrMatch[0]) {
    const attrs = sheetPrMatch[1];
    const selfClose = sheetPrMatch[2] === "/";
    const inner = selfClose ? "" : sheetPrMatch[3] ?? "";
    const innerStripped = inner.replace(/<tabColor\b[^/>]*\/?>/g, "");
    const newInner = `${tabColorChild}${innerStripped}`;
    const replacement = newInner === ""
      ? `<sheetPr${attrs}/>`
      : `<sheetPr${attrs}>${newInner}</sheetPr>`;
    return xml.slice(0, sheetPrMatch.index) + replacement + xml.slice(sheetPrMatch.index + sheetPrMatch[0].length);
  }
  if (!tabColorChild) return xml;
  const wsOpen = /<worksheet\b[^>]*>/.exec(xml);
  if (!wsOpen) return xml;
  const insertAt = wsOpen.index + wsOpen[0].length;
  return xml.slice(0, insertAt) + `<sheetPr>${tabColorChild}</sheetPr>` + xml.slice(insertAt);
}

function injectOpaqueTail(xml: string, sheet: Sheet): string {
  let next = xml;
  // sheetProtection lives much earlier in the canonical order; if we
  // captured it, splice it right after `<dimension>` / `<sheetViews>`.
  if (sheet.sheetProtectionXml) {
    next = next.replace(/<sheetProtection\b[^/>]*\/?>/g, "");
    next = spliceAfterSheetViews(next, sheet.sheetProtectionXml);
  }
  // The rest land at the tail in canonical order.
  for (const block of [
    sheet.printOptionsXml,
    sheet.pageMarginsXml,
    sheet.pageSetupXml,
    sheet.headerFooterXml,
    sheet.rowBreaksXml,
    sheet.colBreaksXml,
    sheet.ignoredErrorsXml,
    sheet.legacyDrawingXml,
    sheet.legacyDrawingHFXml,
    sheet.pictureXml,
    sheet.oleObjectsXml,
    sheet.controlsXml,
  ]) {
    if (!block) continue;
    const tag = /^<([A-Za-z]+)/.exec(block)?.[1];
    if (tag) {
      const stripRe = new RegExp(`<${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</${tag}>)`, "g");
      next = next.replace(stripRe, "");
    }
    next = appendBeforeWorksheetClose(next, block);
  }
  return next;
}

function spliceAfterSheetViews(xml: string, block: string): string {
  const svClose = xml.indexOf("</sheetViews>");
  if (svClose !== -1) {
    const insertAt = svClose + "</sheetViews>".length;
    return xml.slice(0, insertAt) + block + xml.slice(insertAt);
  }
  // No sheetViews — fall back to right after `<dimension>` or `<worksheet>`.
  const dimMatch = /<dimension\b[^/>]*\/?>/.exec(xml);
  if (dimMatch) {
    const insertAt = dimMatch.index + dimMatch[0].length;
    return xml.slice(0, insertAt) + block + xml.slice(insertAt);
  }
  return appendBeforeWorksheetClose(xml, block);
}

function spliceBeforeSheetData(xml: string, block: string): string {
  const sd = xml.indexOf("<sheetData");
  if (sd !== -1) {
    return xml.slice(0, sd) + block + xml.slice(sd);
  }
  return appendBeforeWorksheetClose(xml, block);
}

function spliceBeforePageMargins(xml: string, block: string): string {
  const pm = xml.indexOf("<pageMargins");
  if (pm !== -1) {
    return xml.slice(0, pm) + block + xml.slice(pm);
  }
  return appendBeforeWorksheetClose(xml, block);
}

function appendBeforeWorksheetClose(xml: string, block: string): string {
  const closeIdx = xml.lastIndexOf("</worksheet>");
  if (closeIdx === -1) return xml + block;
  return xml.slice(0, closeIdx) + block + xml.slice(closeIdx);
}

function serializeFilterColumn(colId: number, fc: FilterColumn): string {
  switch (fc.kind) {
    case "values": {
      const filters = [...fc.values].map((v) => `<filter val="${escapeXmlAttr(v)}"/>`).join("");
      const blankAttr = fc.blank ? ' blank="1"' : "";
      return `<filterColumn colId="${colId}"><filters${blankAttr}>${filters}</filters></filterColumn>`;
    }
    case "custom": {
      const ops = [serializeCustomOp(fc.op1)];
      if (fc.op2) ops.push(serializeCustomOp(fc.op2));
      const andAttr = fc.combine === "and" ? ' and="1"' : "";
      return `<filterColumn colId="${colId}"><customFilters${andAttr}>${ops.join("")}</customFilters></filterColumn>`;
    }
    case "top10": {
      const topAttr = fc.top ? "" : ' top="0"';
      const pctAttr = fc.percent ? ' percent="1"' : "";
      const filterValAttr = Number.isFinite(fc.filterVal) ? ` filterVal="${fc.filterVal}"` : "";
      return `<filterColumn colId="${colId}"><top10${topAttr}${pctAttr} val="${fc.n}"${filterValAttr}/></filterColumn>`;
    }
    case "dynamic":
      return `<filterColumn colId="${colId}"><dynamicFilter type="${escapeXmlAttr(fc.type)}"/></filterColumn>`;
    case "color": {
      // We round-trip the dxf id when the parser stamped one. For
      // colour filters set in our UI (where we don't author dxfs) the
      // serializer falls back to dxfId="0" — Excel still loads the
      // file but the colour swatch may not match the original.
      const m = /^dxf:(\d+)$/.exec(fc.argb);
      const dxfId = m ? Number(m[1]) : 0;
      const cellAttr = fc.isCellColor ? "" : ' cellColor="0"';
      return `<filterColumn colId="${colId}"><colorFilter dxfId="${dxfId}"${cellAttr}/></filterColumn>`;
    }
  }
}

function serializeCustomOp(op: CustomFilterOp): string {
  return `<customFilter operator="${op.operator}" val="${escapeXmlAttr(op.val)}"/>`;
}

function colToA1(col: number): string {
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * SheetJS's `write_ws_xml_cell` regenerates the `s` (style) attribute
 * from `cell.z` (number-format string) using its own internal cellXfs
 * table — it ignores any `s` we put on the SheetJS cell object. Since
 * we author our own `xl/styles.xml` from the typed `StyleTable`, we
 * post-process the emitted worksheet XML to inject `s="N"` for every
 * typed cell that carries a `styleId`, and to strip any spurious `s`
 * SheetJS may have added for cells we don't style.
 *
 * Cells without `styleId` get any existing `s` attribute removed so
 * the worksheet XML reflects exactly what's in our typed model. The
 * default xf (index 0) is implicit in OOXML and never written.
 */
/**
 * Re-emit `<f>` elements with shared / array formula metadata for
 * cells whose typed `Formula` carries `kind`, `sharedIndex`, or
 * `ref`. SheetJS expands shared formulas to per-cell text on
 * round-trip, which preserves correctness but bloats the file and
 * destroys the source encoding. This pass surgically rewrites the
 * `<f>` element back to its compact form when the typed model
 * remembers it. Cells without metadata (and cells with literal
 * values, no formula) are untouched — SheetJS's emission stays
 * authoritative.
 */
function injectFormulas(xml: string, cells: ReadonlyMap<string, Cell>): string {
  // Build a per-ref index of cells that need formula rewriting.
  const formulaMeta = new Map<
    string,
    { text: string; kind: "shared" | "array"; sharedIndex?: number; ref?: string; isMaster: boolean }
  >();
  for (const cell of cells.values()) {
    const f = cell.formula;
    if (!f || !f.kind || f.kind === "normal") continue;
    formulaMeta.set(formatA1({ row: cell.row, col: cell.col }), {
      text: f.text,
      kind: f.kind,
      ...(f.sharedIndex !== undefined ? { sharedIndex: f.sharedIndex } : {}),
      ...(f.ref ? { ref: f.ref } : {}),
      isMaster: f.isMaster ?? false,
    });
  }
  if (formulaMeta.size === 0) return xml;

  return xml.replace(/<c\b([^>]*?)>([\s\S]*?)<\/c>/g, (whole: string, attrs: string, body: string) => {
    const refMatch = /\br=("|')([^"']+)\1/.exec(attrs);
    if (!refMatch) return whole;
    const ref = refMatch[2]!;
    const meta = formulaMeta.get(ref);
    if (!meta) return whole;
    // Replace any existing `<f>` element. We don't touch `<v>`
    // (cached value) — the SheetJS round-trip already wrote it.
    const fOpen =
      `<f t="${meta.kind}"` +
      (meta.sharedIndex !== undefined ? ` si="${meta.sharedIndex}"` : "") +
      (meta.isMaster && meta.ref ? ` ref="${meta.ref}"` : "");
    const fEl = meta.isMaster ? `${fOpen}>${escapeXmlText(meta.text)}</f>` : `${fOpen}/>`;
    const replaced = body.replace(/<f\b[^>]*?(?:\/>|>[\s\S]*?<\/f>)/, fEl);
    // If the source body had no `<f>` (SheetJS dropped it because
    // the cell carried only a value), insert ours before `<v>` so
    // Excel still treats the cell as a formula on re-load.
    if (replaced === body && !/<f\b/.test(body)) {
      const vIdx = body.indexOf("<v");
      if (vIdx === -1) return `<c${attrs}>${fEl}${body}</c>`;
      return `<c${attrs}>${body.slice(0, vIdx)}${fEl}${body.slice(vIdx)}</c>`;
    }
    return `<c${attrs}>${replaced}</c>`;
  });
}

function injectStyleIds(xml: string, cells: ReadonlyMap<string, Cell>): string {
  const styleByRef = new Map<string, number>();
  for (const cell of cells.values()) {
    if (cell.styleId !== undefined && cell.styleId !== 0) {
      styleByRef.set(formatA1({ row: cell.row, col: cell.col }), cell.styleId);
    }
  }

  return xml.replace(/<c\b([^/>]*?)(\/?)>/g, (_match, attrs: string, selfClose: string) => {
    const refMatch = /\br=("|')([^"']+)\1/.exec(attrs);
    if (!refMatch) return `<c${attrs}${selfClose}>`;
    const ref = refMatch[2];
    const stripped = attrs.replace(/\s+s=("|')[^"']*\1/, "");
    const styleId = styleByRef.get(ref);
    if (styleId === undefined) {
      return `<c${stripped}${selfClose}>`;
    }
    const newAttrs = `${stripped} s="${styleId}"`;
    return `<c${newAttrs}${selfClose}>`;
  });
}

/**
 * Ensure `[Content_Types].xml` carries an `<Override>` entry for every
 * worksheet currently in the model. We only re-emit the part when an
 * entry was actually added; an unchanged content-types file is left
 * byte-identical to keep the round-trip oracle honest.
 */
function rewriteContentTypes(
  workbook: XlsxWorkbook,
  container: ooxml.OoxmlContainer,
  removedSheetParts: ReadonlySet<string>,
  removedMediaParts: ReadonlySet<string>
): void {
  const ct = ooxml.ContentTypes.load(container);
  let mutated = false;
  for (const removed of removedSheetParts) {
    const partName = `/${removed}`;
    if (ct.hasOverride(partName)) {
      ct.removeOverride(partName);
      mutated = true;
    }
  }
  for (const sheet of workbook.sheets) {
    if (sheet.kind !== "worksheet") continue;
    const partName = `/${sheet.partPath}`;
    if (!ct.hasOverride(partName)) {
      ct.addOverride(partName, WORKSHEET_CONTENT_TYPE);
      mutated = true;
    }
    if (sheet.commentsPartPath) {
      const commentsPartName = `/${sheet.commentsPartPath}`;
      if (!ct.hasOverride(commentsPartName)) {
        ct.addOverride(commentsPartName, COMMENTS_CONTENT_TYPE);
        mutated = true;
      }
    }
    const hasDrawables = sheet.images.length > 0 || sheet.charts.length > 0;
    if (sheet.drawingPartPath && hasDrawables) {
      const drawingPartName = `/${sheet.drawingPartPath}`;
      if (!ct.hasOverride(drawingPartName)) {
        ct.addOverride(drawingPartName, DRAWING_CONTENT_TYPE);
        mutated = true;
      }
    }
  }

  // Drop overrides for drawing parts no sheet still references.
  const liveDrawingPartNames = new Set<string>();
  for (const sheet of workbook.sheets) {
    if (sheet.drawingPartPath && (sheet.images.length > 0 || sheet.charts.length > 0)) {
      liveDrawingPartNames.add(`/${sheet.drawingPartPath}`);
    }
  }
  for (const o of [...ct.overrides]) {
    if (o.contentType === DRAWING_CONTENT_TYPE && !liveDrawingPartNames.has(o.partName)) {
      ct.removeOverride(o.partName);
      mutated = true;
    }
  }

  // Chart content-type overrides: one per `xl/charts/chartN.xml`
  // currently in the container. We register against the container
  // rather than the workbook model because chart parts are derived
  // state — the model doesn't carry their paths.
  const liveChartPartNames = new Set<string>();
  for (const path of container.parts.keys()) {
    if (path.startsWith("xl/charts/") && /chart\d+\.xml$/.test(path)) {
      liveChartPartNames.add(`/${path}`);
    }
  }
  for (const partName of liveChartPartNames) {
    if (!ct.hasOverride(partName)) {
      ct.addOverride(partName, CHART_CONTENT_TYPE);
      mutated = true;
    }
  }
  for (const o of [...ct.overrides]) {
    if (o.contentType === CHART_CONTENT_TYPE && !liveChartPartNames.has(o.partName)) {
      ct.removeOverride(o.partName);
      mutated = true;
    }
  }

  // Media uses `<Default Extension>` entries (one per file extension).
  // Walk the live media set; add an extension default for each
  // distinct file type.
  const liveExtensions = new Set<string>();
  for (const blob of workbook.images.values()) {
    const ext = EXTENSION_BY_CONTENT_TYPE[blob.contentType];
    if (ext) liveExtensions.add(ext);
  }
  for (const ext of liveExtensions) {
    if (!ct.hasDefault(ext)) {
      ct.addDefault(ext, mimeForExtension(ext));
      mutated = true;
    }
  }
  void removedMediaParts;

  if (mutated) ct.writeBack(container);
}

function mimeForExtension(ext: string): string {
  if (ext === "png") return "image/png";
  if (ext === "jpeg" || ext === "jpg") return "image/jpeg";
  if (ext === "gif") return "image/gif";
  return `image/${ext}`;
}

/**
 * Re-emit `xl/comments{N}.xml` for every dirty comments part. The
 * owning sheet is found by matching `commentsPartPath`. Brand-new
 * comments parts (the first comment on a sheet) reach this path
 * because the handler set `commentsPartPath` and dirtied it before
 * serialize ran.
 */
function rewriteDirtyComments(
  workbook: XlsxWorkbook,
  container: ooxml.OoxmlContainer,
  paths: ReadonlySet<string>
): void {
  for (const path of paths) {
    const sheet = workbook.sheets.find((s) => s.commentsPartPath === path);
    if (!sheet) {
      throw new XlsxSerializeError("container-failed", `dirty comments path ${path} not owned by any sheet`, {
        partPath: path,
      });
    }
    const xml = serializeCommentsPart(sheet.commentAuthors, sheet.comments);
    container.writeText(path, xml);
  }
}

/**
 * Rewrite per-sheet rels parts (`xl/worksheets/_rels/sheetN.xml.rels`)
 * that have been dirtied — currently this is exclusively driven by
 * `xlsx:add-comment`, which needs to add a `comments` relationship.
 *
 * Strategy: load the existing rels (so any pre-existing hyperlink /
 * vmlDrawing / drawing rels are preserved verbatim), drop any
 * relationship typed `…/relationships/comments`, then re-add a single
 * comments relationship pointing at the sheet's `commentsPartPath`.
 *
 * VML drawings — the legacy `<v:shape>` markup that anchors classic
 * notes visually in Excel — are deferred to P1. Without VML the
 * comment round-trips in the data layer but won't render with a
 * pinned position in Excel; the headless P0 surface accepts that.
 */
function rewriteDirtySheetRels(
  workbook: XlsxWorkbook,
  container: ooxml.OoxmlContainer,
  paths: ReadonlySet<string>
): void {
  for (const relsPath of paths) {
    const sheet = workbook.sheets.find((s) => ooxml.RelationshipGraph.relsPathFor(s.partPath) === relsPath);
    if (!sheet) {
      throw new XlsxSerializeError(
        "container-failed",
        `dirty sheet rels path ${relsPath} not owned by any sheet`,
        { partPath: relsPath }
      );
    }
    const graph = ooxml.RelationshipGraph.loadFor(container, sheet.partPath);
    for (const r of [...graph.relationships]) {
      if (r.type === COMMENTS_REL_TYPE) graph.remove(r.id);
    }
    if (sheet.commentsPartPath) {
      const target = relsRelativeTarget(sheet.partPath, sheet.commentsPartPath);
      graph.add({ type: COMMENTS_REL_TYPE, target });
    }
    graph.writeBack(container);
  }
}

/**
 * Compute a rels `Target` from `sheet.partPath` to `commentsPartPath`,
 * matching how Excel writes the relationship: `../comments1.xml` when
 * the sheet lives under `xl/worksheets/`.
 */
function relsRelativeTarget(ownerPartPath: string, targetPath: string): string {
  const ownerDir = ownerPartPath.includes("/") ? ownerPartPath.slice(0, ownerPartPath.lastIndexOf("/")) : "";
  const ownerSegs = ownerDir ? ownerDir.split("/") : [];
  const targetSegs = targetPath.split("/");

  let common = 0;
  while (
    common < ownerSegs.length &&
    common < targetSegs.length &&
    ownerSegs[common] === targetSegs[common]
  ) {
    common++;
  }
  const ups = ownerSegs.length - common;
  const rest = targetSegs.slice(common).join("/");
  return ups > 0 ? `${"../".repeat(ups)}${rest}` : rest;
}

/**
 * Ensure `xl/_rels/workbook.xml.rels` carries a `<Relationship>` for
 * every worksheet. New sheets get a freshly minted `rId`; existing
 * relationships (sharedStrings, styles, theme, calcChain, …) are
 * preserved verbatim. As with content-types we only write back when
 * something was actually added.
 */
function rewriteWorkbookRels(
  workbook: XlsxWorkbook,
  container: ooxml.OoxmlContainer,
  removedSheetParts: ReadonlySet<string>
): void {
  const rels = ooxml.RelationshipGraph.loadFor(container, WORKBOOK_PART);

  let mutated = false;
  if (removedSheetParts.size > 0) {
    const removedTargets = new Set<string>();
    for (const path of removedSheetParts) removedTargets.add(workbookRelTarget(path));
    const orphans = rels.relationships.filter(
      (r) => r.type === WORKSHEET_REL_TYPE && removedTargets.has(normalizeRelTarget(r.target))
    );
    for (const r of orphans) {
      rels.remove(r.id);
      mutated = true;
    }
  }

  const covered = new Set<string>();
  for (const r of rels.relationships) covered.add(normalizeRelTarget(r.target));

  for (const sheet of workbook.sheets) {
    if (sheet.kind !== "worksheet") continue;
    const target = workbookRelTarget(sheet.partPath);
    if (!covered.has(target)) {
      rels.add({ type: WORKSHEET_REL_TYPE, target });
      covered.add(target);
      mutated = true;
    }
  }
  if (mutated) rels.writeBack(container);
}

/**
 * Drop sheet parts (and their `_rels/` sidecars) listed in
 * `dirty.removedSheetParts`. The workbook-rels and content-types
 * pruning happens in their respective rewrite passes; this function
 * only touches the part files themselves.
 */
function dropRemovedSheetParts(
  container: ooxml.OoxmlContainer,
  removedSheetParts: ReadonlySet<string>
): void {
  for (const path of removedSheetParts) {
    const relsPath = ooxml.RelationshipGraph.relsPathFor(path);
    // Best-effort drop of any drawing part the removed sheet owned —
    // its `_rels` would otherwise be orphaned. Media bytes are GC'd
    // at command-handler time when no sheet still references them, so
    // they live in `removedMediaParts` instead of being dropped here.
    if (container.has(relsPath)) {
      const drawingRel = ooxml.RelationshipGraph.loadFor(container, path).relationships.find(
        (r) => r.type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
      );
      if (drawingRel) {
        const drawingPath = drawingRel.target.startsWith("/")
          ? drawingRel.target.slice(1)
          : joinRel(path, drawingRel.target);
        if (container.has(drawingPath)) container.removePart(drawingPath);
        const drawingRelsPath = ooxml.RelationshipGraph.relsPathFor(drawingPath);
        if (container.has(drawingRelsPath)) container.removePart(drawingRelsPath);
      }
      container.removePart(relsPath);
    }
    if (container.has(path)) container.removePart(path);
  }
}

function joinRel(ownerPartPath: string, target: string): string {
  const ownerDir = ownerPartPath.includes("/") ? ownerPartPath.slice(0, ownerPartPath.lastIndexOf("/")) : "";
  const segments = (ownerDir ? ownerDir.split("/") : []).concat(target.split("/"));
  const stack: string[] = [];
  for (const seg of segments) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") stack.pop();
    else stack.push(seg);
  }
  return stack.join("/");
}

/**
 * Re-emit just the `<sheets>` element of `xl/workbook.xml` from the
 * typed `workbook.sheets` array. Every other byte of the workbook
 * part (namespaces, comments, attribute order on `<workbook>` itself,
 * `<bookViews>`, `<definedNames>`, `<calcPr>`, …) is left untouched
 * via a string-level splice.
 *
 * The rebuild covers the three commands that currently dirty the
 * workbook part: rename, add-sheet, and (eventually) reorder. The
 * `r:id` for each sheet is looked up via the workbook rels, which
 * the caller is responsible for refreshing first when new sheets
 * have been added.
 */
function rewriteWorkbookSheets(workbook: XlsxWorkbook, container: ooxml.OoxmlContainer): void {
  if (!container.has(WORKBOOK_PART)) return;
  const xml = container.readText(WORKBOOK_PART);

  const rels = ooxml.RelationshipGraph.loadFor(container, WORKBOOK_PART);
  const ridByTarget = new Map<string, string>();
  for (const r of rels.relationships) {
    ridByTarget.set(normalizeRelTarget(r.target), r.id);
  }

  const sheetEntries: string[] = [];
  for (const sheet of workbook.sheets) {
    const target = workbookRelTarget(sheet.partPath);
    const rid = ridByTarget.get(target);
    if (!rid) {
      throw new XlsxSerializeError(
        "workbook-failed",
        `No workbook relationship for sheet "${sheet.name}" (target=${target}); did the rels rewrite run?`
      );
    }
    const stateAttr = sheet.state !== "visible" ? ` state="${sheet.state}"` : "";
    sheetEntries.push(
      `<sheet name="${escapeXmlAttr(sheet.name)}" sheetId="${escapeXmlAttr(sheet.sheetId)}"${stateAttr} r:id="${escapeXmlAttr(rid)}"/>`
    );
  }
  const newSheetsBlock = `<sheets>${sheetEntries.join("")}</sheets>`;

  const sheetsRe = /<sheets\b[^>]*?(?:\/>|>[\s\S]*?<\/sheets>)/;
  const match = sheetsRe.exec(xml);
  if (!match) {
    throw new XlsxSerializeError("workbook-failed", "Could not locate <sheets> block in xl/workbook.xml");
  }
  let next = xml.slice(0, match.index) + newSheetsBlock + xml.slice(match.index + match[0].length);

  // C12 — Re-emit the `<definedNames>` block from the typed
  // `workbook.definedNames` array. We always emit it on workbook
  // dirty (even when empty) so that removing the last name in the
  // session also drops the block from the OOXML.
  next = injectDefinedNames(next, workbook);

  // Phase 6 — workbook protection. The verbatim block lives on
  // `workbookProtectionXml`; mutated by `xlsx:set-workbook-protection`.
  // Always rewrite on workbook dirty so a protection that was cleared
  // also drops out of the on-disk part.
  next = injectWorkbookProtection(next, workbook);

  // Phase 4c — calcPr. Mutated by `xlsx:set-calc-mode` (changes
  // `@calcMode`, `@iterate`, `@calcOnSave` attributes verbatim).
  next = injectCalcPr(next, workbook);

  container.writeText(WORKBOOK_PART, next);
}

function injectCalcPr(xml: string, workbook: XlsxWorkbook): string {
  const block = workbook.calcPrXml ?? "";
  const re = /<calcPr\b[^>]*?(?:\/>|>[\s\S]*?<\/calcPr>)/;
  const m = re.exec(xml);
  if (m) {
    return xml.slice(0, m.index) + block + xml.slice(m.index + m[0].length);
  }
  if (!block) return xml;
  // Canonical position: right before `</workbook>`.
  return xml.replace(/<\/workbook>/, `${block}</workbook>`);
}

function injectWorkbookProtection(xml: string, workbook: XlsxWorkbook): string {
  const block = workbook.workbookProtectionXml ?? "";
  const re = /<workbookProtection\b[^>]*?(?:\/>|>[\s\S]*?<\/workbookProtection>)/;
  const m = re.exec(xml);
  if (m) {
    return xml.slice(0, m.index) + block + xml.slice(m.index + m[0].length);
  }
  if (!block) return xml;
  // Canonical position: after `<fileVersion>`/`<workbookPr>` and
  // before `<bookViews>` / `<sheets>`. Splice immediately before
  // `<bookViews>` when present, otherwise before `<sheets>`.
  const bookViewsRe = /<bookViews\b/;
  const bv = bookViewsRe.exec(xml);
  if (bv) return xml.slice(0, bv.index) + block + xml.slice(bv.index);
  const sheetsRe = /<sheets\b/;
  const sm = sheetsRe.exec(xml);
  if (sm) return xml.slice(0, sm.index) + block + xml.slice(sm.index);
  return xml.replace(/<\/workbook>/, `${block}</workbook>`);
}

function injectDefinedNames(xml: string, workbook: XlsxWorkbook): string {
  // Build the new block (or empty string when there are no names).
  let block = "";
  if (workbook.definedNames.length > 0) {
    const nameByIndex = new Map<string, number>();
    workbook.sheets.forEach((s, i) => nameByIndex.set(s.name, i));
    const entries: string[] = [];
    for (const dn of workbook.definedNames) {
      const attrs: string[] = [`name="${escapeXmlAttr(dn.name)}"`];
      if (dn.scope) {
        const idx = nameByIndex.get(dn.scope);
        if (idx !== undefined) attrs.push(`localSheetId="${idx}"`);
      }
      if (dn.hidden) attrs.push(`hidden="1"`);
      if (dn.comment) attrs.push(`comment="${escapeXmlAttr(dn.comment)}"`);
      entries.push(`<definedName ${attrs.join(" ")}>${escapeXmlText(dn.refersTo)}</definedName>`);
    }
    block = `<definedNames>${entries.join("")}</definedNames>`;
  }

  const dnRe = /<definedNames\b[^>]*?(?:\/>|>[\s\S]*?<\/definedNames>)/;
  const m = dnRe.exec(xml);
  if (m) {
    return xml.slice(0, m.index) + block + xml.slice(m.index + m[0].length);
  }
  if (!block) return xml;
  // Excel's canonical position is between </sheets> and <calcPr> (or
  // before </workbook> when neither exists). Splice immediately after
  // the closing </sheets>.
  const sheetsCloseRe = /<\/sheets>/;
  const sm = sheetsCloseRe.exec(xml);
  if (sm) {
    const idx = sm.index + sm[0].length;
    return xml.slice(0, idx) + block + xml.slice(idx);
  }
  // Fallback — shove it right before </workbook>.
  return xml.replace(/<\/workbook>/, `${block}</workbook>`);
}

/**
 * Re-emit `xl/styles.xml` from the typed style table. Round-trip is
 * **semantic** (not byte-identical): attribute order can drift, but
 * re-parsing the output yields a structurally equivalent table.
 */
function rewriteStylesXml(workbook: XlsxWorkbook, master: ooxml.OoxmlContainer): void {
  const xml = serializeStylesXml(workbook.styles);
  if (master.has(STYLES_PART)) {
    master.writeText(STYLES_PART, xml);
  } else {
    master.writeText(STYLES_PART, xml);
  }
}

function workbookRelTarget(partPath: string): string {
  return partPath.startsWith("xl/") ? partPath.slice(3) : partPath;
}

function normalizeRelTarget(target: string): string {
  return target.startsWith("/") ? target.slice(1) : target;
}

function escapeXmlAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/**
 * Element-text escape. Quotes/apostrophes don't need encoding inside
 * text nodes per the XML spec; only `&`, `<`, and `>` do.
 */
function escapeXmlText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
