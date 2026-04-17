import { ooxml } from "@officeai/core";
import * as XLSX from "xlsx";
import type { Sheet, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { XlsxSerializeError } from "./errors.js";
import { syncSheetToSheetJS } from "./sheet-sync.js";

const WORKBOOK_PART = "xl/workbook.xml";

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
    dirty.styles ||
    dirty.contentTypes ||
    dirty.rels ||
    dirty.comments.size > 0 ||
    dirty.threadedComments.size > 0 ||
    dirty.sheetRels.size > 0;

  if (unsupportedDirty) {
    throw new XlsxSerializeError(
      "container-failed",
      "Phase 5 serializer supports only dirty `sheets` + `workbook`; sst/styles/rels/comments rewrites land in later phases"
    );
  }

  if (dirtySheetPaths.size > 0) {
    await rewriteDirtySheets(snapshot, container, dirtySheetPaths);
  }

  if (dirty.workbook) {
    rewriteWorkbookSheetNames(snapshot.root, container);
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
    const newBytes = emittedContainer.readBytes(emittedPath);
    master.writeBytes(path, newBytes);
  }
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

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function escapeXmlAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
