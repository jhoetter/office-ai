import { ooxml } from "@officeai/core";
import * as XLSX from "xlsx";
import { EXTENSION_BY_CONTENT_TYPE } from "../model/drawings.js";
import { formatA1 } from "../model/refs.js";
import type {
  AutoFilter,
  Cell,
  CustomFilterOp,
  DataValidation,
  FilterColumn,
  FreezePanes,
  Sheet,
  XlsxSnapshot,
  XlsxWorkbook,
} from "../model/types.js";
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
  if (dirty.drawings.size > 0) {
    rewriteDirtyDrawings(snapshot.root, container, dirty.drawings, drawingRidByPath);
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
    xml = injectHiddenRows(xml, sheet.hiddenRows);
    xml = injectAutoFilter(xml, sheet.autoFilter);
    xml = injectFreezePanes(xml, sheet.freeze);
    xml = injectConditionalFormats(xml, sheet.opaqueConditionalFormats);
    xml = injectDataValidations(xml, sheet.dataValidations, sheet.opaqueDataValidations);
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
  drawingRidByPath: Map<string, string | null>
): void {
  for (const sheetPartPath of paths) {
    const sheet = workbook.sheets.find((s) => s.partPath === sheetPartPath);
    if (!sheet) continue;

    const sheetRels = ooxml.RelationshipGraph.loadFor(container, sheetPartPath);

    if (sheet.images.length === 0) {
      // No images left → drop the drawing part + rels and clear the
      // sheet's drawing relationship.
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
    const { graph: drawingRels, embedRidByMediaPath } = buildDrawingRels(
      drawingPartPath,
      sheet.images,
      workbook.images
    );
    const xml = serializeDrawingPart(sheet.images, embedRidByMediaPath);
    container.writeText(drawingPartPath, xml);
    drawingRels.writeBack(container);

    const rid = upsertSheetDrawingRel(sheetRels, sheetPartPath, drawingPartPath);
    sheetRels.writeBack(container);
    drawingRidByPath.set(sheetPartPath, rid);
  }
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
function injectConditionalFormats(xml: string, opaqueBlocks: ReadonlyArray<string>): string {
  // First, drop any pre-existing CF blocks so we don't double-emit
  // when SheetJS already echoed them through.
  const stripped = xml.replace(
    /<conditionalFormatting\b[^>]*?(?:\/>|>[\s\S]*?<\/conditionalFormatting>)/g,
    ""
  );
  if (opaqueBlocks.length === 0) return stripped;
  const closeIdx = stripped.lastIndexOf("</worksheet>");
  const block = opaqueBlocks.join("");
  if (closeIdx === -1) return stripped + block;
  return stripped.slice(0, closeIdx) + block + stripped.slice(closeIdx);
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

function injectFreezePanes(xml: string, freeze: FreezePanes | undefined): string {
  const next = xml.replace(/<sheetViews\b[^>]*>[\s\S]*?<\/sheetViews>/g, "");
  const viewsBlock = renderSheetViews(freeze);
  if (!viewsBlock) return next;

  // Excel orders `<sheetViews>` immediately after `<dimension>` (or
  // first thing in `<worksheet>` if `<dimension>` isn't present).
  const dimMatch = /<dimension\b[^/>]*\/?>(?:<\/dimension>)?/.exec(next);
  if (dimMatch) {
    const insertAt = dimMatch.index + dimMatch[0].length;
    return next.slice(0, insertAt) + viewsBlock + next.slice(insertAt);
  }
  const wsMatch = /<worksheet\b[^>]*>/.exec(next);
  if (wsMatch) {
    const insertAt = wsMatch.index + wsMatch[0].length;
    return next.slice(0, insertAt) + viewsBlock + next.slice(insertAt);
  }
  return next + viewsBlock;
}

function renderSheetViews(freeze: FreezePanes | undefined): string {
  if (!freeze || (freeze.rows <= 0 && freeze.cols <= 0)) {
    // Nothing to add. Excel happily reads a worksheet without a
    // `<sheetViews>` block (it falls back to defaults).
    return "";
  }
  const xSplit = Math.max(0, Math.floor(freeze.cols));
  const ySplit = Math.max(0, Math.floor(freeze.rows));
  const topLeftRow = ySplit + 1;
  const topLeftCol = colToA1(xSplit);
  const topLeft = `${topLeftCol}${topLeftRow}`;
  const activePane = xSplit > 0 && ySplit > 0 ? "bottomRight" : xSplit > 0 ? "topRight" : "bottomLeft";
  const xAttr = xSplit > 0 ? ` xSplit="${xSplit}"` : "";
  const yAttr = ySplit > 0 ? ` ySplit="${ySplit}"` : "";
  return (
    `<sheetViews>` +
    `<sheetView tabSelected="1" workbookViewId="0">` +
    `<pane${xAttr}${yAttr} topLeftCell="${topLeft}" activePane="${activePane}" state="frozen"/>` +
    `</sheetView>` +
    `</sheetViews>`
  );
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
    if (sheet.drawingPartPath && sheet.images.length > 0) {
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
    if (sheet.drawingPartPath && sheet.images.length > 0) {
      liveDrawingPartNames.add(`/${sheet.drawingPartPath}`);
    }
  }
  for (const o of [...ct.overrides]) {
    if (o.contentType === DRAWING_CONTENT_TYPE && !liveDrawingPartNames.has(o.partName)) {
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

  container.writeText(WORKBOOK_PART, next);
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
