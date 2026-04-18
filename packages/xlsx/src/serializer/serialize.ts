import { ooxml } from "@officeai/core";
import * as XLSX from "xlsx";
import { formatA1 } from "../model/refs.js";
import type { Cell, Sheet, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { XlsxSerializeError } from "./errors.js";
import { syncSheetToSheetJS } from "./sheet-sync.js";
import { serializeStylesXml } from "./styles.js";

const WORKBOOK_PART = "xl/workbook.xml";
const STYLES_PART = "xl/styles.xml";

/**
 * Serialize an `XlsxSnapshot` back to bytes.
 *
 * Phase 5 contract:
 *   - No dirty flags set → byte-content-identical re-emit, identical
 *     to Phase 4. Untouched workbooks always round-trip exactly.
 *   - Dirty sheets → for each dirty sheet, sync the typed cells +
 *     merges back onto the SheetJS WorkSheet, then ask SheetJS to
 *     emit a single-sheet workbook for that sheet, and substitute the
 *     emitted `xl/worksheets/sheetN.xml` into the master container.
 *     Other parts (workbook.xml, sst, styles, opaque parts, …) stay
 *     byte-identical.
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
  const unsupportedDirty =
    dirty.sharedStrings ||
    dirty.contentTypes ||
    dirty.rels ||
    dirty.comments.size > 0 ||
    dirty.threadedComments.size > 0 ||
    dirty.sheetRels.size > 0;

  if (unsupportedDirty) {
    throw new XlsxSerializeError(
      "container-failed",
      "Phase 5 serializer supports only dirty `sheets` + `workbook` + `styles`; sst/rels/comments rewrites land in later phases"
    );
  }

  if (dirtySheetPaths.size > 0) {
    await rewriteDirtySheets(snapshot, container, dirtySheetPaths);
  }

  if (dirty.workbook) {
    rewriteWorkbookSheetNames(snapshot.root, container);
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
 * Phase 5 surgical patch: only the `<sheet name="...">` attributes
 * inside `<sheets>` change. We re-emit the workbook XML by
 * search-and-replace over the original text so every other byte
 * (namespaces, comments, ordering, attribute order, custom views,
 * defined names, calc props, …) stays byte-identical. Sheet names are
 * uniquely identified by their `r:id`, which is stable across renames.
 */
function rewriteWorkbookSheetNames(workbook: XlsxWorkbook, master: ooxml.OoxmlContainer): void {
  if (!master.has(WORKBOOK_PART)) return;
  let xml = master.readText(WORKBOOK_PART);

  const tree = ooxml.parseXml(xml);
  const root = (tree as unknown[]).map((n) => ooxml.asElement(n)).find((el) => el?.tag === "workbook");
  if (!root) return;
  const sheetsEl = ooxml.findChild(root.children, "sheets");
  if (!sheetsEl) return;

  const sheetEls = ooxml.filterChildren(sheetsEl.children, "sheet");
  const renamesByRid = new Map<string, { oldName: string; newName: string }>();
  for (const sheet of workbook.sheets) {
    const matching = sheetEls.find((el) => el.attrs.sheetId === sheet.sheetId);
    if (!matching) continue;
    const oldName = matching.attrs.name;
    if (oldName !== undefined && oldName !== sheet.name) {
      const rid = matching.attrs["r:id"] ?? matching.attrs["r:Id"] ?? "";
      renamesByRid.set(rid, { oldName, newName: sheet.name });
    }
  }
  if (renamesByRid.size === 0) return;

  for (const [rid, { oldName, newName }] of renamesByRid) {
    const ridEsc = escapeRegex(rid);
    const oldEsc = escapeRegex(escapeXmlAttr(oldName));
    const re = new RegExp(`(<sheet\\b[^>]*\\bname=")${oldEsc}("[^>]*\\br:id="${ridEsc}")`, "g");
    const re2 = new RegExp(`(<sheet\\b[^>]*\\br:id="${ridEsc}"[^>]*\\bname=")${oldEsc}(")`, "g");
    const replaced = xml.replace(re, `$1${escapeXmlAttr(newName)}$2`);
    xml = replaced.replace(re2, `$1${escapeXmlAttr(newName)}$2`);
  }

  master.writeText(WORKBOOK_PART, xml);
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

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function escapeXmlAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
