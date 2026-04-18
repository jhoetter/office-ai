import { ooxml } from "@officeai/core";
import * as XLSX from "xlsx";
import { formatA1 } from "../model/refs.js";
import type { Cell, Sheet, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { serializeCommentsPart } from "./comments.js";
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

  if (dirtySheetPaths.size > 0) {
    await rewriteDirtySheets(snapshot, container, dirtySheetPaths);
  }

  if (dirty.comments.size > 0) {
    rewriteDirtyComments(snapshot.root, container, dirty.comments);
  }

  if (dirty.sheetRels.size > 0) {
    rewriteDirtySheetRels(snapshot.root, container, dirty.sheetRels);
  }

  if (dirty.contentTypes) {
    rewriteContentTypes(snapshot.root, container);
  }

  if (dirty.rels) {
    rewriteWorkbookRels(snapshot.root, container);
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
  dirtySheetPaths: ReadonlySet<string>
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
    master.writeText(path, xml);
  }
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
function rewriteContentTypes(workbook: XlsxWorkbook, container: ooxml.OoxmlContainer): void {
  const ct = ooxml.ContentTypes.load(container);
  let mutated = false;
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
  }
  if (mutated) ct.writeBack(container);
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
function rewriteWorkbookRels(workbook: XlsxWorkbook, container: ooxml.OoxmlContainer): void {
  const rels = ooxml.RelationshipGraph.loadFor(container, WORKBOOK_PART);
  const covered = new Set<string>();
  for (const r of rels.relationships) covered.add(normalizeRelTarget(r.target));

  let mutated = false;
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
  const next = xml.slice(0, match.index) + newSheetsBlock + xml.slice(match.index + match[0].length);
  container.writeText(WORKBOOK_PART, next);
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
