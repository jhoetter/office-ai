import { defaultIdMinter, ooxml, sha256Hex, type IdMinter } from "@officeai/core";
import * as XLSX from "@e965/xlsx";
import { cellKey, formatA1, parseA1, parseRange } from "../model/refs.js";
import { defaultStyleTable, type StyleTable } from "../model/style-table.js";
import {
  emptyDirty,
  type AutoFilter,
  type Cell,
  type CellErrorCode,
  type CellValue,
  type Comment,
  type CustomFilterOp,
  type DataValidation,
  type DynamicFilterType,
  type FilterColumn,
  type Formula,
  type FreezePanes,
  type MergedCell,
  type OpaquePart,
  type Sheet,
  type XlsxSnapshot,
  type XlsxWorkbook,
} from "../model/types.js";
import type { ImageBlob } from "../model/index.js";
import type { NodeId } from "@officeai/core";
import { parseCommentsPart } from "./comments.js";
import { resolveDrawings } from "./drawings.js";
import { XlsxParseError } from "./errors.js";
import { discoverPivotParts } from "./pivot-tables.js";
import { parseStylesXml } from "./styles.js";

const WORKBOOK_PART = "xl/workbook.xml";
const WORKBOOK_RELS_PART = "xl/_rels/workbook.xml.rels";
const CONTENT_TYPES_PART = "[Content_Types].xml";

const DOC_REL_TYPES = {
  worksheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
  chartsheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
  dialogsheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
} as const;

const COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

const WORKBOOK_ROOT_ATTR_PREFIXES = ["xmlns", "xmlns:", "mc:Ignorable", "xml:space"];

/**
 * Parts the typed model (will) own. Anything outside this set is parked
 * in `opaqueParts` for byte-clean round-trip. Keys are exact part paths
 * or path prefixes (suffix-matched by `pathStartsWith`).
 */
const MODELED_EXACT_PATHS = new Set<string>([
  WORKBOOK_PART,
  WORKBOOK_RELS_PART,
  CONTENT_TYPES_PART,
  "_rels/.rels",
  "xl/sharedStrings.xml",
  "xl/styles.xml",
]);

const MODELED_PREFIXES = ["xl/worksheets/", "xl/_rels/", "xl/comments"];

function isModeledPath(path: string): boolean {
  if (MODELED_EXACT_PATHS.has(path)) return true;
  return MODELED_PREFIXES.some((p) => path.startsWith(p));
}

export interface ParseOptions {
  readonly idMinter?: IdMinter;
}

export async function parseXlsx(
  input: ArrayBuffer | Uint8Array,
  opts: ParseOptions = {}
): Promise<XlsxSnapshot> {
  let container: ooxml.OoxmlContainer;
  try {
    container = await ooxml.OoxmlContainer.load(input);
  } catch (err) {
    throw new XlsxParseError("zip-corruption", "Failed to read XLSX as a zip archive", {
      cause: err,
    });
  }

  if (!container.has(WORKBOOK_PART)) {
    throw new XlsxParseError("missing-workbook-part", `Missing required part: ${WORKBOOK_PART}`, {
      partPath: WORKBOOK_PART,
    });
  }
  if (!container.has(CONTENT_TYPES_PART)) {
    throw new XlsxParseError("missing-content-types", `Missing required part: ${CONTENT_TYPES_PART}`, {
      partPath: CONTENT_TYPES_PART,
    });
  }

  const mintNodeId: IdMinter = opts.idMinter ?? defaultIdMinter;

  const {
    workbookRootAttrs,
    date1904,
    sheetEntries,
    definedNames: rawDefinedNames,
  } = parseWorkbookXml(container);

  let sheetjsBook: XLSX.WorkBook;
  try {
    const buf = input instanceof Uint8Array ? input : new Uint8Array(input);
    sheetjsBook = XLSX.read(buf, {
      type: "array",
      dense: true,
      cellFormula: true,
      cellStyles: true,
      cellNF: true,
      cellDates: false,
      sheetStubs: true,
      bookFiles: false,
      bookVBA: true,
      xlfn: true,
    });
  } catch (err) {
    throw new XlsxParseError("sheetjs-failure", "SheetJS failed to read the workbook", { cause: err });
  }

  const workbookRels = ooxml.RelationshipGraph.loadFor(container, WORKBOOK_PART);
  const images = new Map<string, ImageBlob>();
  const sheets: Sheet[] = sheetEntries.map((entry, index) =>
    resolveSheet(entry, index, workbookRels, container, sheetjsBook, mintNodeId, images)
  );

  const contentTypes = ooxml.ContentTypes.load(container);
  const ctMap = new Map<string, string>();
  for (const o of contentTypes.overrides) {
    const partName = o.partName.startsWith("/") ? o.partName.slice(1) : o.partName;
    ctMap.set(partName, o.contentType);
  }

  // Drawing parts and media we successfully modeled — exclude from
  // opaqueParts so we don't double-store them, and so commands that
  // remove an image can drop the bytes from the package without a
  // stale opaque copy resurrecting them.
  const modeledDrawingParts = new Set<string>();
  for (const sheet of sheets) {
    if (sheet.drawingPartPath) {
      modeledDrawingParts.add(sheet.drawingPartPath);
      modeledDrawingParts.add(ooxml.RelationshipGraph.relsPathFor(sheet.drawingPartPath));
    }
  }
  const modeledMediaParts = new Set(images.keys());

  // F1 Phase 1 — pivot tables + caches lift from the catch-all
  // `opaqueParts` bucket into typed `pivotTables` / `pivotCaches`
  // slots on the workbook. The bytes still travel verbatim via the
  // typed record's `raw` field; the serializer re-emits them
  // byte-identical when no typed pivot edit has occurred (always in
  // Phase 1). See `spec/xlsx/pivot-tables.md`.
  const pivot = discoverPivotParts(container, ctMap, sheets);

  const opaqueParts = new Map<string, OpaquePart>();
  const partHashes: Record<string, string> = {};
  for (const [path, part] of container.parts) {
    const hash = sha256Hex(part.bytes);
    partHashes[path] = hash;
    if (isModeledPath(path)) continue;
    if (modeledDrawingParts.has(path)) continue;
    if (modeledMediaParts.has(path)) continue;
    if (pivot.modeledPaths.has(path)) continue;
    opaqueParts.set(path, {
      path,
      bytes: part.bytes,
      contentType: ctMap.get(path),
      hash,
    });
  }

  const styles: StyleTable = container.has("xl/styles.xml")
    ? parseStylesXml(container.readText("xl/styles.xml"))
    : defaultStyleTable();

  // C12 — Resolve OOXML `localSheetId` indices (0-based) into sheet
  // names so commands can reason about scope without re-walking the
  // sheet array. Out-of-range indices fall back to workbook scope.
  const definedNames = rawDefinedNames.map((n) => {
    const scope =
      n.localSheetId !== undefined && n.localSheetId >= 0 && n.localSheetId < sheets.length
        ? sheets[n.localSheetId]!.name
        : undefined;
    const out: import("../model/types.js").DefinedName = {
      id: `dn-${n.name}-${scope ?? "wb"}`,
      name: n.name,
      refersTo: n.refersTo,
      ...(scope ? { scope } : {}),
      ...(n.comment ? { comment: n.comment } : {}),
      ...(n.hidden ? { hidden: true } : {}),
    };
    return out;
  });

  const workbookXmlText = container.readText(WORKBOOK_PART);
  const workbookProtectionXml = extractFirstBlock(workbookXmlText, "workbookProtection");
  const calcPrXml = extractFirstBlock(workbookXmlText, "calcPr");

  const workbook: XlsxWorkbook = {
    id: mintNodeId(),
    sheets,
    partHashes,
    opaqueParts,
    date1904,
    workbookRootAttrs,
    ...(workbookProtectionXml ? { workbookProtectionXml } : {}),
    ...(calcPrXml ? { calcPrXml } : {}),
    styles,
    sheetjs: sheetjsBook,
    images,
    definedNames,
    pivotTables: pivot.pivotTables,
    pivotCaches: pivot.pivotCaches,
  };

  return {
    format: "xlsx",
    revision: 0,
    root: workbook,
    partHashes,
    container,
    dirty: emptyDirty(),
  };
}

interface RawSheetEntry {
  readonly name: string;
  readonly sheetId: string;
  readonly rId: string;
  readonly state: "visible" | "hidden" | "veryHidden";
}

interface RawDefinedName {
  readonly name: string;
  readonly refersTo: string;
  readonly localSheetId?: number;
  readonly comment?: string;
  readonly hidden?: boolean;
}

function parseWorkbookXml(container: ooxml.OoxmlContainer): {
  workbookRootAttrs: Record<string, string>;
  date1904: boolean;
  sheetEntries: RawSheetEntry[];
  definedNames: RawDefinedName[];
} {
  const xml = container.readText(WORKBOOK_PART);
  let tree: unknown;
  try {
    tree = ooxml.parseXml(xml);
  } catch (err) {
    throw new XlsxParseError("invalid-xml", `Failed to parse ${WORKBOOK_PART}`, {
      partPath: WORKBOOK_PART,
      cause: err,
    });
  }

  if (!Array.isArray(tree)) {
    throw new XlsxParseError("invalid-workbook", `${WORKBOOK_PART} root must be an element list`, {
      partPath: WORKBOOK_PART,
    });
  }

  const root = (tree as unknown[]).map((n) => ooxml.asElement(n)).find((el) => el?.tag === "workbook");
  if (!root) {
    throw new XlsxParseError("invalid-workbook", `${WORKBOOK_PART} missing <workbook> root`, {
      partPath: WORKBOOK_PART,
    });
  }

  const workbookRootAttrs: Record<string, string> = {};
  for (const [k, v] of Object.entries(root.attrs)) {
    if (k === "xmlns" || WORKBOOK_ROOT_ATTR_PREFIXES.some((p) => k === p || k.startsWith(p))) {
      workbookRootAttrs[k] = v;
    } else {
      workbookRootAttrs[k] = v;
    }
  }

  const wbPr = ooxml.findChild(root.children, "workbookPr");
  const date1904 = wbPr?.attrs.date1904 === "1" || wbPr?.attrs.date1904 === "true";

  const sheetsEl = ooxml.findChild(root.children, "sheets");
  const sheetEntries: RawSheetEntry[] = [];
  if (sheetsEl) {
    for (const child of ooxml.filterChildren(sheetsEl.children, "sheet")) {
      const name = child.attrs.name;
      const sheetId = child.attrs.sheetId;
      const rId = child.attrs["r:id"] ?? child.attrs["r:Id"];
      if (!name || !sheetId || !rId) {
        throw new XlsxParseError(
          "invalid-workbook",
          `<sheet> entry missing required attribute (name/sheetId/r:id) in ${WORKBOOK_PART}`,
          { partPath: WORKBOOK_PART }
        );
      }
      const stateRaw = child.attrs.state;
      const state: RawSheetEntry["state"] =
        stateRaw === "hidden" || stateRaw === "veryHidden" ? stateRaw : "visible";
      sheetEntries.push({ name, sheetId, rId, state });
    }
  }

  const definedNamesEl = ooxml.findChild(root.children, "definedNames");
  const definedNames: RawDefinedName[] = [];
  if (definedNamesEl) {
    for (const child of ooxml.filterChildren(definedNamesEl.children, "definedName")) {
      const name = child.attrs.name;
      if (!name) continue;
      // refersTo lives in the element text. Strip leading `=` per
      // OOXML convention; Excel writes both with and without it.
      const text = ooxml.getTextContent(child.entry).trim();
      const refersTo = text.startsWith("=") ? text.slice(1) : text;
      const lsRaw = child.attrs.localSheetId;
      const ls = lsRaw !== undefined ? Number.parseInt(lsRaw, 10) : Number.NaN;
      const hidden = child.attrs.hidden === "1" || child.attrs.hidden === "true";
      const comment = child.attrs.comment;
      const entry: RawDefinedName = {
        name,
        refersTo,
        ...(Number.isFinite(ls) ? { localSheetId: ls } : {}),
        ...(comment ? { comment } : {}),
        ...(hidden ? { hidden: true } : {}),
      };
      definedNames.push(entry);
    }
  }

  return { workbookRootAttrs, date1904, sheetEntries, definedNames };
}

function resolveSheet(
  entry: RawSheetEntry,
  index: number,
  workbookRels: ooxml.RelationshipGraph,
  container: ooxml.OoxmlContainer,
  sheetjsBook: XLSX.WorkBook,
  mintNodeId: IdMinter,
  images: Map<string, ImageBlob>
): Sheet {
  const rel = workbookRels.byId(entry.rId);
  if (!rel) {
    throw new XlsxParseError(
      "missing-sheet-target",
      `Workbook references rId="${entry.rId}" but no matching relationship in xl/_rels/workbook.xml.rels`,
      { partPath: WORKBOOK_RELS_PART }
    );
  }
  const partPath = resolveTargetPath(WORKBOOK_PART, rel.target);
  if (!container.has(partPath)) {
    throw new XlsxParseError(
      "missing-sheet-target",
      `Sheet "${entry.name}" target ${partPath} missing from container`,
      { partPath }
    );
  }
  const relsPath = ooxml.RelationshipGraph.relsPathFor(partPath);
  const relsPartPath = container.has(relsPath) ? relsPath : undefined;

  const kind: Sheet["kind"] = rel.type === DOC_REL_TYPES.worksheet ? "worksheet" : "non-worksheet";

  const ws = sheetjsBook.Sheets[entry.name];
  const sheetXml = kind === "worksheet" && container.has(partPath) ? container.readText(partPath) : undefined;
  const styleIdByRef = sheetXml ? extractStyleIdsFromXml(sheetXml) : undefined;
  // `cells` is reassigned below to fold in shared/array formula
  // metadata; `merges` stays referentially stable.
  // eslint-disable-next-line prefer-const
  let { cells, merges } =
    kind === "worksheet" && ws ? extractCellsAndMerges(ws, styleIdByRef) : EMPTY_CELL_DATA;

  const { commentsPartPath, comments, commentAuthors } = resolveComments(container, partPath);

  const { autoFilter, hiddenRows, rowHeights } = sheetXml
    ? extractAutoFilterAndHiddenRows(sheetXml)
    : { autoFilter: undefined, hiddenRows: new Set<number>(), rowHeights: new Map<number, number>() };

  const { columnWidths, hiddenCols } = sheetXml
    ? extractColumnWidthsAndHidden(sheetXml)
    : { columnWidths: new Map<number, number>(), hiddenCols: new Set<number>() };

  const sheetDefaults = sheetXml ? extractSheetFormatPr(sheetXml) : {};

  const freeze = sheetXml ? extractFreezePanes(sheetXml) : undefined;
  const opaqueConditionalFormats = sheetXml ? extractConditionalFormatBlocks(sheetXml) : [];
  const dvParse = sheetXml
    ? extractDataValidations(sheetXml)
    : { typed: [] as ReadonlyArray<DataValidation>, opaque: undefined };

  const drawings =
    kind === "worksheet"
      ? resolveDrawings(container, partPath, mintNodeId, images)
      : { images: [] as ReadonlyArray<import("../model/index.js").SheetImage>, mediaBlobs: [] };
  for (const blob of drawings.mediaBlobs) {
    if (!images.has(blob.partPath)) images.set(blob.partPath, blob);
  }

  // C14 — Hydrate Excel Tables (`<tableParts>` referencing
  // `xl/tables/tableN.xml`). Tables we successfully parse get their
  // part path recorded so the serializer can mark it as modeled
  // (and skip the byte-clean opaqueParts copy).
  const tables =
    kind === "worksheet" && sheetXml ? resolveTables(container, partPath, sheetXml, mintNodeId) : [];

  const opaqueSheetParts = sheetXml ? extractOpaqueSheetParts(sheetXml) : EMPTY_OPAQUE_SHEET_PARTS;

  // Shared / array formula metadata. SheetJS expands shared
  // formulas to per-cell text on parse, which loses the source
  // grouping. We re-scan the worksheet XML directly to recover the
  // `t="shared"` / `t="array"` attributes plus their `si` and
  // `ref` so the serializer can re-emit the compact form. Cells
  // that lack metadata are normal formulas and round-trip via
  // SheetJS unchanged.
  if (kind === "worksheet" && sheetXml) {
    const meta = extractFormulaMetadata(sheetXml);
    if (meta.size > 0) {
      const next = new Map(cells);
      // Map of follower-ref → master text so we can backfill formula
      // text on followers that SheetJS dropped (because their `<f>`
      // was self-closing with no body).
      const masterTextBySi = new Map<number, string>();
      for (const [, m] of meta) {
        if (m.isMaster && m.si !== undefined) {
          // We need the master cell's text — pull it after the loop;
          // cache the marker for now.
          masterTextBySi.set(m.si, "");
        }
      }
      for (const [a1Ref, m] of meta) {
        const addr = parseA1(a1Ref);
        const key = cellKey(addr.row, addr.col);
        const existing = next.get(key);
        if (m.isMaster && m.si !== undefined && existing?.formula) {
          masterTextBySi.set(m.si, existing.formula.text);
        }
      }
      for (const [a1Ref, m] of meta) {
        const addr = parseA1(a1Ref);
        const key = cellKey(addr.row, addr.col);
        const existing = next.get(key);
        if (!existing) continue;
        const fallbackText =
          existing.formula?.text ?? (m.si !== undefined ? (masterTextBySi.get(m.si) ?? "") : "");
        next.set(key, {
          ...existing,
          formula: {
            text: fallbackText,
            ...(m.kind ? { kind: m.kind } : {}),
            ...(m.si !== undefined ? { sharedIndex: m.si } : {}),
            ...(m.ref ? { ref: m.ref } : {}),
            ...(m.isMaster ? { isMaster: true } : {}),
          },
        });
      }
      cells = next;
    }
  }

  return {
    id: mintNodeId(),
    sheetId: entry.sheetId,
    name: entry.name,
    index,
    state: entry.state,
    kind,
    partPath,
    ...(relsPartPath ? { relsPartPath } : {}),
    cells,
    merges,
    comments,
    commentAuthors,
    columnWidths,
    rowHeights,
    hiddenCols,
    hiddenRows,
    ...(sheetDefaults.defaultColWidthPx !== undefined
      ? { defaultColWidthPx: sheetDefaults.defaultColWidthPx }
      : {}),
    ...(sheetDefaults.defaultRowHeightPx !== undefined
      ? { defaultRowHeightPx: sheetDefaults.defaultRowHeightPx }
      : {}),
    images: drawings.images,
    conditionalFormats: [],
    opaqueConditionalFormats,
    dataValidations: dvParse.typed,
    ...(dvParse.opaque ? { opaqueDataValidations: dvParse.opaque } : {}),
    ...("drawingPartPath" in drawings && drawings.drawingPartPath
      ? { drawingPartPath: drawings.drawingPartPath }
      : {}),
    ...(autoFilter ? { autoFilter } : {}),
    ...(freeze ? { freeze } : {}),
    ...(commentsPartPath ? { commentsPartPath } : {}),
    tables,
    charts: [],
    ...(opaqueSheetParts.hyperlinksXml ? { hyperlinksXml: opaqueSheetParts.hyperlinksXml } : {}),
    ...(opaqueSheetParts.tablePartsXml ? { tablePartsXml: opaqueSheetParts.tablePartsXml } : {}),
    ...(opaqueSheetParts.colsXml ? { colsXml: opaqueSheetParts.colsXml } : {}),
    ...(opaqueSheetParts.sheetViewsXml ? { sheetViewsXml: opaqueSheetParts.sheetViewsXml } : {}),
    ...(opaqueSheetParts.sheetProtectionXml
      ? { sheetProtectionXml: opaqueSheetParts.sheetProtectionXml }
      : {}),
    ...(opaqueSheetParts.pageMarginsXml ? { pageMarginsXml: opaqueSheetParts.pageMarginsXml } : {}),
    ...(opaqueSheetParts.pageSetupXml ? { pageSetupXml: opaqueSheetParts.pageSetupXml } : {}),
    ...(opaqueSheetParts.printOptionsXml ? { printOptionsXml: opaqueSheetParts.printOptionsXml } : {}),
    ...(opaqueSheetParts.headerFooterXml ? { headerFooterXml: opaqueSheetParts.headerFooterXml } : {}),
    ...(opaqueSheetParts.rowBreaksXml ? { rowBreaksXml: opaqueSheetParts.rowBreaksXml } : {}),
    ...(opaqueSheetParts.colBreaksXml ? { colBreaksXml: opaqueSheetParts.colBreaksXml } : {}),
    ...(opaqueSheetParts.ignoredErrorsXml ? { ignoredErrorsXml: opaqueSheetParts.ignoredErrorsXml } : {}),
    ...(opaqueSheetParts.legacyDrawingXml ? { legacyDrawingXml: opaqueSheetParts.legacyDrawingXml } : {}),
    ...(opaqueSheetParts.legacyDrawingHFXml
      ? { legacyDrawingHFXml: opaqueSheetParts.legacyDrawingHFXml }
      : {}),
    ...(opaqueSheetParts.pictureXml ? { pictureXml: opaqueSheetParts.pictureXml } : {}),
    ...(opaqueSheetParts.oleObjectsXml ? { oleObjectsXml: opaqueSheetParts.oleObjectsXml } : {}),
    ...(opaqueSheetParts.controlsXml ? { controlsXml: opaqueSheetParts.controlsXml } : {}),
  };
}

/**
 * Capture per-worksheet OOXML blocks that the typed model does not
 * own but that a dirty rewrite via SheetJS would otherwise drop.
 *
 * Each capture is a verbatim string slice of the input XML so the
 * serializer can re-inject it into the regenerated worksheet XML
 * with no shape transformations. The injection order matters
 * (see `serializer/serialize.ts#injectOpaqueSheetParts`); the
 * extractor is order-agnostic.
 */
interface OpaqueSheetParts {
  readonly hyperlinksXml?: string;
  readonly tablePartsXml?: string;
  readonly colsXml?: string;
  readonly sheetViewsXml?: string;
  readonly sheetProtectionXml?: string;
  readonly pageMarginsXml?: string;
  readonly pageSetupXml?: string;
  readonly printOptionsXml?: string;
  readonly headerFooterXml?: string;
  readonly rowBreaksXml?: string;
  readonly colBreaksXml?: string;
  readonly ignoredErrorsXml?: string;
  readonly legacyDrawingXml?: string;
  readonly legacyDrawingHFXml?: string;
  readonly pictureXml?: string;
  readonly oleObjectsXml?: string;
  readonly controlsXml?: string;
}

const EMPTY_OPAQUE_SHEET_PARTS: OpaqueSheetParts = {};

function extractFirstBlock(xml: string, tag: string): string | undefined {
  const re = new RegExp(`<${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</${tag}>)`);
  const m = re.exec(xml);
  return m ? m[0] : undefined;
}

export function extractOpaqueSheetParts(xml: string): OpaqueSheetParts {
  return {
    ...optional("hyperlinksXml", extractFirstBlock(xml, "hyperlinks")),
    ...optional("tablePartsXml", extractFirstBlock(xml, "tableParts")),
    ...optional("colsXml", extractFirstBlock(xml, "cols")),
    ...optional("sheetViewsXml", extractFirstBlock(xml, "sheetViews")),
    ...optional("sheetProtectionXml", extractFirstBlock(xml, "sheetProtection")),
    ...optional("pageMarginsXml", extractFirstBlock(xml, "pageMargins")),
    ...optional("pageSetupXml", extractFirstBlock(xml, "pageSetup")),
    ...optional("printOptionsXml", extractFirstBlock(xml, "printOptions")),
    ...optional("headerFooterXml", extractFirstBlock(xml, "headerFooter")),
    ...optional("rowBreaksXml", extractFirstBlock(xml, "rowBreaks")),
    ...optional("colBreaksXml", extractFirstBlock(xml, "colBreaks")),
    ...optional("ignoredErrorsXml", extractFirstBlock(xml, "ignoredErrors")),
    ...optional("legacyDrawingXml", extractFirstBlock(xml, "legacyDrawing")),
    ...optional("legacyDrawingHFXml", extractFirstBlock(xml, "legacyDrawingHF")),
    ...optional("pictureXml", extractFirstBlock(xml, "picture")),
    ...optional("oleObjectsXml", extractFirstBlock(xml, "oleObjects")),
    ...optional("controlsXml", extractFirstBlock(xml, "controls")),
  };
}

function optional<K extends string>(key: K, value: string | undefined): Partial<Record<K, string>> {
  return value ? ({ [key]: value } as Record<K, string>) : {};
}

interface FormulaMeta {
  readonly kind?: "shared" | "array";
  readonly si?: number;
  readonly ref?: string;
  readonly isMaster?: boolean;
}

/**
 * Scan worksheet XML for `<c r="…"><f t="shared|array" si="…"
 * ref="…">…</f></c>` patterns and produce a per-cell map of the
 * shared / array formula metadata. SheetJS exposes the resolved
 * formula text on every cell but loses the encoding metadata, so
 * we recover it directly from the XML to drive the serializer's
 * `<f>` rewrite.
 *
 * Master vs follower:
 * - Master cells are the ones whose `<f>` element carries the
 *   formula body (and, for shared groups, the `ref` attribute).
 * - Follower cells reference the master via the same `si` and have
 *   either a self-closing `<f t="shared" si=…/>` or a redundant
 *   body that Excel ignores.
 */
export function extractFormulaMetadata(xml: string): Map<string, FormulaMeta> {
  const out = new Map<string, FormulaMeta>();
  const cellRe = /<c\b([^>]*?)>([\s\S]*?)<\/c>/g;
  let cm: RegExpExecArray | null;
  while ((cm = cellRe.exec(xml))) {
    const attrs = cm[1] ?? "";
    const body = cm[2] ?? "";
    const refMatch = /\br=("|')([^"']+)\1/.exec(attrs);
    if (!refMatch) continue;
    const ref = refMatch[2]!;
    const fMatch = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(body);
    if (!fMatch) continue;
    const fAttrs = fMatch[1] ?? "";
    const fBody = fMatch[2];
    const tMatch = /\bt=("|')([^"']+)\1/.exec(fAttrs);
    const kind = tMatch?.[2] === "shared" ? "shared" : tMatch?.[2] === "array" ? "array" : undefined;
    if (!kind) continue;
    const siMatch = /\bsi=("|')(\d+)\1/.exec(fAttrs);
    const refMatchInF = /\bref=("|')([^"']+)\1/.exec(fAttrs);
    const isMaster = (fBody !== undefined && fBody.length > 0) || refMatchInF !== null;
    const meta: FormulaMeta = {
      kind,
      ...(siMatch ? { si: Number.parseInt(siMatch[2]!, 10) } : {}),
      ...(refMatchInF ? { ref: refMatchInF[2]! } : {}),
      ...(isMaster ? { isMaster: true } : {}),
    };
    out.set(ref, meta);
  }
  return out;
}

/**
 * C14 — Resolve `<tableParts>` from the sheet XML into typed
 * {@link TableDef} records. We:
 *
 * 1. Walk `<tableParts><tablePart r:id="…"/>` to enumerate the rels.
 * 2. Look up each rel in the sheet's rels part to find the table
 *    part path (`xl/tables/tableN.xml`).
 * 3. Parse the table XML for `id`, `name`, `displayName`, `ref`,
 *    `headerRowCount`, `totalsRowCount`, `<tableColumns>`,
 *    `<tableStyleInfo>`, and `<autoFilter>`.
 * 4. Keep the original XML so we can re-emit verbatim when the table
 *    isn't dirty.
 *
 * Anything that fails to parse is silently skipped — the underlying
 * part still lands in `opaqueParts` via the catch-all in `parseXlsx`,
 * so the file round-trips correctly even if we couldn't hydrate it.
 */
function resolveTables(
  container: ooxml.OoxmlContainer,
  sheetPartPath: string,
  sheetXml: string,
  mintNodeId: () => NodeId
): ReadonlyArray<import("../model/types.js").TableDef> {
  const tablePartRe = /<tablePart\s+[^>]*?r:id="([^"]+)"[^>]*\/?>/g;
  const relIds: string[] = [];
  let m: RegExpExecArray | null;
  while ((m = tablePartRe.exec(sheetXml)) !== null) {
    relIds.push(m[1]!);
  }
  if (relIds.length === 0) return [];

  const sheetRels = ooxml.RelationshipGraph.loadFor(container, sheetPartPath);

  const out: import("../model/types.js").TableDef[] = [];
  for (const relId of relIds) {
    const rel = sheetRels.byId(relId);
    if (!rel) continue;
    const partPath = resolveRelativePath(sheetPartPath, rel.target);
    if (!container.has(partPath)) continue;
    const xml = container.readText(partPath);
    const def = parseTableXml(xml, partPath, relId, mintNodeId);
    if (def) out.push(def);
  }
  return out;
}

/**
 * Resolve `target` (a path inside the package, possibly relative
 * with `../`) against the directory of `basePath`. Mirrors the
 * `xml:base` resolution Excel uses for rels targets.
 */
function resolveRelativePath(basePath: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const baseDir = basePath.includes("/") ? basePath.slice(0, basePath.lastIndexOf("/")) : "";
  const segments = baseDir.split("/").filter(Boolean);
  for (const part of target.split("/")) {
    if (part === "" || part === ".") continue;
    if (part === "..") {
      segments.pop();
    } else {
      segments.push(part);
    }
  }
  return segments.join("/");
}

function parseTableXml(
  xml: string,
  partPath: string,
  relId: string,
  mintNodeId: () => NodeId
): import("../model/types.js").TableDef | undefined {
  const tableMatch = /<table\b([^>]*)>/.exec(xml);
  if (!tableMatch) return undefined;
  const attrs = parseAttrs(tableMatch[1]!);
  const tableId = attrs.id;
  const name = attrs.name;
  const displayName = attrs.displayName ?? name;
  const range = attrs.ref;
  if (!tableId || !name || !range) return undefined;

  const headerRowCount = attrs.headerRowCount !== undefined ? Number.parseInt(attrs.headerRowCount, 10) : 1;
  const totalsRowCount = attrs.totalsRowCount !== undefined ? Number.parseInt(attrs.totalsRowCount, 10) : 0;

  const columnNames: string[] = [];
  const colRe = /<tableColumn\b([^>]*?)\/?>/g;
  let cm: RegExpExecArray | null;
  while ((cm = colRe.exec(xml)) !== null) {
    const ca = parseAttrs(cm[1]!);
    if (ca.name) columnNames.push(decodeXmlAttr(ca.name));
  }

  const styleMatch = /<tableStyleInfo\b[^>]*\/?>/.exec(xml);
  const autoFilterMatch = /<autoFilter\s+[^>]*?ref="([^"]+)"/.exec(xml);

  return {
    id: mintNodeId(),
    tableId,
    name,
    displayName,
    range,
    headerRowCount: Number.isFinite(headerRowCount) ? headerRowCount : 1,
    totalsRowCount: Number.isFinite(totalsRowCount) ? totalsRowCount : 0,
    columnNames,
    ...(styleMatch ? { styleInfoXml: styleMatch[0]! } : {}),
    ...(autoFilterMatch ? { autoFilterRange: autoFilterMatch[1]! } : {}),
    opaqueXml: xml,
    partPath,
    relId,
  };
}

function parseAttrs(input: string): Record<string, string> {
  const out: Record<string, string> = {};
  const re = /([A-Za-z_:][\w:.-]*)="([^"]*)"/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(input)) !== null) {
    out[m[1]!] = m[2]!;
  }
  return out;
}

function decodeXmlAttr(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

/**
 * Parse `<sheetView><pane xSplit="N" ySplit="M" state="frozen"/></sheetView>`
 * into a {@link FreezePanes} record. We only honour `state="frozen"`
 * (and the implicit "frozen" when no `state` is present yet a split
 * is); split-bar (`state="split"`) and `frozenSplit` are preserved as
 * opaque XML — the parser returns `undefined` so the serializer leaves
 * the original `<sheetView>` block untouched.
 *
 * Excel emits `xSplit` (cols) and `ySplit` (rows). Either may be
 * absent (= 0). When both are 0 there's no freeze.
 */
/**
 * Capture every `<conditionalFormatting …>…</conditionalFormatting>`
 * block from the worksheet XML, verbatim. We treat them as opaque
 * strings so the serializer can re-inject them on dirty round-trip
 * without needing to model every cfRule shape (data bars, icon
 * sets, color scales, formula rules, etc.). Typed CF authoring
 * lives alongside this in {@link Sheet.conditionalFormats} (C10).
 */
export function extractConditionalFormatBlocks(xml: string): ReadonlyArray<string> {
  const out: string[] = [];
  const re = /<conditionalFormatting\b[^>]*?(?:\/>|>[\s\S]*?<\/conditionalFormatting>)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    out.push(m[0]);
  }
  return out;
}

/**
 * C11 — Pull `<dataValidations>` content out of a worksheet part.
 *
 * Returns:
 *   - `typed`  — list-kind validations parsed into model rules
 *   - `opaque` — the raw `<dataValidations>…</dataValidations>` block
 *                with **typed** `list` rules stripped, so the
 *                serializer can re-emit non-list rules verbatim
 *                next to a freshly-emitted typed block. `undefined`
 *                when nothing remains after stripping.
 *
 * SheetJS itself doesn't surface dataValidations, so we lift them
 * straight from the XML. The regex tolerates self-closing entries
 * (most files) and full element form (rare authoring tools).
 */
export function extractDataValidations(xml: string): {
  typed: ReadonlyArray<DataValidation>;
  opaque: string | undefined;
} {
  const blockRe = /<dataValidations\b([^>]*)>([\s\S]*?)<\/dataValidations>/;
  const block = blockRe.exec(xml);
  if (!block) return { typed: [], opaque: undefined };
  const inner = block[2] ?? "";
  const ruleRe = /<dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/dataValidation>)/g;
  const typed: DataValidation[] = [];
  const remaining: string[] = [];
  let nextId = 1;
  let m: RegExpExecArray | null;
  while ((m = ruleRe.exec(inner)) !== null) {
    const attrs = m[1] ?? "";
    const body = m[2] ?? "";
    const typeMatch = /\btype=("|')([^"']+)\1/.exec(attrs);
    const sqrefMatch = /\bsqref=("|')([^"']+)\1/.exec(attrs);
    if (typeMatch && typeMatch[2] === "list" && sqrefMatch) {
      // <formula1> body holds the source. Quoted-literal lists look
      // like `"Yes,No"`; references look like `Sheet1!$A$1:$A$5`.
      const f1 = /<formula1>([\s\S]*?)<\/formula1>/.exec(body);
      const raw = f1?.[1]?.trim() ?? "";
      const isQuoted = raw.startsWith('"') && raw.endsWith('"');
      const source = isQuoted ? raw.slice(1, -1) : raw;
      const showDropDown = !/\bshowDropDown=("|')1\1/.test(attrs);
      const stopOnInvalid = !/\berrorStyle=("|')(warning|information)\1/.test(attrs);
      const allowBlank = /\ballowBlank=("|')1\1/.test(attrs);
      typed.push({
        kind: "list",
        id: `dv-import-${nextId++}`,
        range: sqrefMatch[2]!,
        source,
        formula: !isQuoted,
        showDropDown,
        stopOnInvalid,
        allowBlank,
      });
    } else {
      remaining.push(m[0]);
    }
  }
  if (remaining.length === 0) return { typed, opaque: undefined };
  // Re-wrap the remaining (non-list) rules in a fresh
  // <dataValidations> element so the serializer can drop it back in
  // verbatim. We don't try to preserve the original count attribute
  // because Excel ignores it on read.
  const opaque = `<dataValidations>${remaining.join("")}</dataValidations>`;
  return { typed, opaque };
}

export function extractFreezePanes(xml: string): FreezePanes | undefined {
  const m = /<pane\b([^/>]*)\/?>/.exec(xml);
  if (!m) return undefined;
  const attrs = m[1] ?? "";
  const stateMatch = /\bstate=("|')([^"']+)\1/.exec(attrs);
  if (stateMatch && stateMatch[2] !== "frozen") return undefined;
  const xMatch = /\bxSplit=("|')([^"']+)\1/.exec(attrs);
  const yMatch = /\bySplit=("|')([^"']+)\1/.exec(attrs);
  const cols = xMatch ? Math.max(0, Math.floor(Number(xMatch[2]))) : 0;
  const rows = yMatch ? Math.max(0, Math.floor(Number(yMatch[2]))) : 0;
  if (!Number.isFinite(cols) || !Number.isFinite(rows)) return undefined;
  if (cols === 0 && rows === 0) return undefined;
  return { rows, cols };
}

function resolveComments(
  container: ooxml.OoxmlContainer,
  sheetPartPath: string
): {
  commentsPartPath?: string;
  comments: ReadonlyArray<Comment>;
  commentAuthors: ReadonlyArray<string>;
} {
  const relsPath = ooxml.RelationshipGraph.relsPathFor(sheetPartPath);
  if (!container.has(relsPath)) return { comments: [], commentAuthors: [] };

  const rels = ooxml.RelationshipGraph.loadFor(container, sheetPartPath);
  const commentRel = rels.relationships.find((r) => r.type === COMMENTS_REL_TYPE);
  if (!commentRel) return { comments: [], commentAuthors: [] };

  const commentsPartPath = resolveTargetPath(sheetPartPath, commentRel.target);
  if (!container.has(commentsPartPath)) return { comments: [], commentAuthors: [] };

  const xml = container.readText(commentsPartPath);
  const { authors, comments } = parseCommentsPart(xml, commentsPartPath);
  return { commentsPartPath, comments, commentAuthors: authors };
}

const EMPTY_CELL_DATA: { cells: ReadonlyMap<string, Cell>; merges: ReadonlyArray<MergedCell> } = {
  cells: new Map(),
  merges: [],
};

interface SheetJSCellLike {
  readonly t?: string;
  readonly v?: unknown;
  readonly f?: string;
  /** Style index when the workbook was loaded with `cellStyles: true`. */
  readonly s?: number;
}

interface SheetJSWorksheetLike {
  readonly "!data"?: ReadonlyArray<ReadonlyArray<SheetJSCellLike | undefined>>;
  readonly "!merges"?: ReadonlyArray<{ s: { r: number; c: number }; e: { r: number; c: number } }>;
}

function extractCellsAndMerges(
  ws: XLSX.WorkSheet,
  styleIdByRef?: ReadonlyMap<string, number>
): {
  cells: ReadonlyMap<string, Cell>;
  merges: ReadonlyArray<MergedCell>;
} {
  const dense = ws as unknown as SheetJSWorksheetLike;
  const cells = new Map<string, Cell>();

  const data = dense["!data"];
  if (data) {
    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const raw = row[c];
        if (!raw) continue;
        const styleId = styleIdByRef?.get(formatA1({ row: r, col: c }));
        const cell = sheetjsCellToTyped(raw, r, c, styleId);
        if (cell) cells.set(cellKey(r, c), cell);
      }
    }
  }

  const merges: MergedCell[] = (dense["!merges"] ?? []).map((m) => ({
    r1: m.s.r,
    c1: m.s.c,
    r2: m.e.r,
    c2: m.e.c,
  }));

  return { cells, merges };
}

function sheetjsCellToTyped(
  raw: SheetJSCellLike,
  row: number,
  col: number,
  styleIdFromXml: number | undefined
): Cell | null {
  const t = raw.t;
  let value: CellValue;

  switch (t) {
    case "n":
      value = typeof raw.v === "number" ? raw.v : Number(raw.v ?? 0);
      break;
    case "s":
    case "str":
      value = String(raw.v ?? "");
      break;
    case "b":
      value = Boolean(raw.v);
      break;
    case "d":
      value = typeof raw.v === "number" ? raw.v : String(raw.v ?? "");
      break;
    case "e":
      value = mapErrorCode(raw.v);
      break;
    case "z":
    case undefined:
      if (raw.v === undefined && !raw.f) return null;
      value = raw.v === undefined ? null : (raw.v as CellValue);
      break;
    default:
      value = raw.v === undefined ? null : (raw.v as CellValue);
  }

  const formula: Formula | undefined = raw.f ? { text: raw.f } : undefined;
  const styleId = styleIdFromXml ?? (typeof raw.s === "number" ? raw.s : undefined);
  if (value === null && !formula && styleId === undefined) return null;

  const cell: Cell = {
    row,
    col,
    value,
    ...(formula ? { formula } : {}),
    ...(styleId !== undefined ? { styleId } : {}),
  };
  return cell;
}

/**
 * Extract per-cell style indices (`s` attribute on `<c>`) directly from
 * the worksheet XML. SheetJS replaces `cell.s` with a resolved fill
 * object when loaded with `cellStyles: true`, so the original numeric
 * index is lost from the dense store. We need the index to drive our
 * typed `StyleTable`; SheetJS's resolved fills/fonts are not enough
 * because we own the styles part end-to-end.
 *
 * The regex tolerates attribute order, single or double quotes, and
 * self-closing `<c .../>` cells.
 */
function extractStyleIdsFromXml(xml: string): Map<string, number> {
  const result = new Map<string, number>();
  const re = /<c\b([^/>]*?)\/?>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    const attrs = m[1];
    const refMatch = /\br=("|')([^"']+)\1/.exec(attrs);
    if (!refMatch) continue;
    const sMatch = /\bs=("|')([^"']+)\1/.exec(attrs);
    if (!sMatch) continue;
    const idx = Number(sMatch[2]);
    if (!Number.isInteger(idx) || idx < 0) continue;
    result.set(refMatch[2], idx);
  }
  return result;
}

/**
 * Extract `<autoFilter>` and `<row r="N" hidden="1">` from a worksheet
 * XML string. Regex-based to match the same posture as
 * {@link extractStyleIdsFromXml} — full XML parsing would pull in
 * another dependency just to read two patterns.
 *
 * Excel allows at most one `<autoFilter>` per worksheet; we only
 * honour the first match.
 */
export function extractAutoFilterAndHiddenRows(xml: string): {
  autoFilter: AutoFilter | undefined;
  hiddenRows: Set<number>;
  rowHeights: Map<number, number>;
} {
  const hiddenRows = new Set<number>();
  const rowHeights = new Map<number, number>();
  // `<row …>` carries `r`, `hidden`, `ht`, `customHeight` — we only
  // honour `ht` when `customHeight="1"` (matches Excel: a `ht`
  // without `customHeight` is a hint, not an override). Heights are
  // points; convert to CSS px at 96 dpi.
  const rowRe = /<row\b([^/>]*)\/?>/g;
  let rm: RegExpExecArray | null;
  while ((rm = rowRe.exec(xml)) !== null) {
    const attrs = rm[1] ?? "";
    const rMatch = /\br=("|')([^"']+)\1/.exec(attrs);
    if (!rMatch) continue;
    const rowNum = Number(rMatch[2]);
    if (!Number.isInteger(rowNum) || rowNum < 1) continue;
    const rowIdx = rowNum - 1;
    const hMatch = /\bhidden=("|')([^"']+)\1/.exec(attrs);
    if (hMatch) {
      const hVal = hMatch[2];
      if (hVal === "1" || hVal === "true") hiddenRows.add(rowIdx);
    }
    const customHeight = /\bcustomHeight=("|')(1|true)\1/.test(attrs);
    if (customHeight) {
      const htMatch = /\bht=("|')([^"']+)\1/.exec(attrs);
      if (htMatch) {
        const ptVal = Number(htMatch[2]);
        if (Number.isFinite(ptVal) && ptVal >= 0) {
          rowHeights.set(rowIdx, ptToPx(ptVal));
        }
      }
    }
  }

  const afMatch = /<autoFilter\b([^>]*)(?:\/>|>([\s\S]*?)<\/autoFilter>)/.exec(xml);
  let autoFilter: AutoFilter | undefined;
  if (afMatch) {
    const headerAttrs = afMatch[1] ?? "";
    const body = afMatch[2] ?? "";
    const refMatch = /\bref=("|')([^"']+)\1/.exec(headerAttrs);
    if (refMatch) {
      try {
        const range = parseRange(refMatch[2]!);
        const columns = parseFilterColumns(body);
        autoFilter = {
          range: {
            r1: range.start.row,
            c1: range.start.col,
            r2: range.end.row,
            c2: range.end.col,
          },
          columns,
        };
      } catch {
        // Malformed ref — skip silently rather than fail the whole parse.
      }
    }
  }

  return { autoFilter, hiddenRows, rowHeights };
}

/**
 * Excel column widths are stored in "character units" — the width of
 * a `0` in the workbook's default font. The OOXML reference formula
 * for converting back to CSS pixels at the workbook's max-digit
 * width is involved enough to need the theme font metrics; for our
 * render-time approximation `Math.round(width * 7 + 5)` is within
 * a pixel of Excel for the typical Calibri 11pt default and avoids
 * a theme lookup. Hidden columns and explicit `width="0"` collapse
 * to 0 px so the renderer's hidden-column rule paints them as
 * zero-width.
 */
function charWidthToPx(width: number): number {
  if (!Number.isFinite(width) || width <= 0) return 0;
  return Math.round(width * 7 + 5);
}

/** Excel point sizes → CSS px at 96 dpi. */
function ptToPx(pt: number): number {
  return Math.round(pt * (96 / 72));
}

/**
 * Parse `<cols><col min="N" max="M" width="X" hidden="…"/></cols>`
 * into a per-column-index map of pixel widths plus a set of hidden
 * column indices. Each `<col>` covers the inclusive 1-based range
 * `[min, max]`; we expand it to flat 0-based entries so the
 * renderer doesn't have to know about Excel's range encoding — it
 * just looks up `columnWidths.get(c)` / `hiddenCols.has(c)`.
 *
 * Hidden columns are also entered into `columnWidths` (set to 0) so
 * the Grid's axis-index treats them uniformly with non-hidden
 * overrides.
 */
export function extractColumnWidthsAndHidden(xml: string): {
  columnWidths: Map<number, number>;
  hiddenCols: Set<number>;
} {
  const columnWidths = new Map<number, number>();
  const hiddenCols = new Set<number>();
  const blockMatch = /<cols\b[^>]*>([\s\S]*?)<\/cols>/.exec(xml);
  if (!blockMatch) return { columnWidths, hiddenCols };
  const inner = blockMatch[1] ?? "";
  const colRe = /<col\b([^/>]*)\/?>/g;
  let cm: RegExpExecArray | null;
  while ((cm = colRe.exec(inner)) !== null) {
    const attrs = cm[1] ?? "";
    const minMatch = /\bmin=("|')([^"']+)\1/.exec(attrs);
    const maxMatch = /\bmax=("|')([^"']+)\1/.exec(attrs);
    if (!minMatch || !maxMatch) continue;
    const min = Number(minMatch[2]);
    const max = Number(maxMatch[2]);
    if (!Number.isInteger(min) || !Number.isInteger(max) || min < 1 || max < min) continue;
    const widthMatch = /\bwidth=("|')([^"']+)\1/.exec(attrs);
    const widthChars = widthMatch ? Number(widthMatch[2]) : Number.NaN;
    const hidden = /\bhidden=("|')(1|true)\1/.test(attrs);
    const widthPx = Number.isFinite(widthChars) ? charWidthToPx(widthChars) : NaN;
    for (let one = min; one <= max; one++) {
      const c = one - 1;
      if (hidden) {
        hiddenCols.add(c);
        columnWidths.set(c, 0);
      } else if (Number.isFinite(widthPx)) {
        columnWidths.set(c, Math.max(0, widthPx));
      }
    }
  }
  return { columnWidths, hiddenCols };
}

/**
 * Parse `<sheetFormatPr defaultColWidth="…" defaultRowHeight="…"/>`
 * into per-sheet pixel defaults. Both attributes are optional; we
 * leave the corresponding field unset when missing so the renderer
 * falls through to its built-in defaults.
 *
 * `baseColWidth` (Excel's "standard" column width) is intentionally
 * ignored here because Excel itself derives `defaultColWidth` from
 * it via a font-metric formula we don't replicate; the explicit
 * `defaultColWidth` is what Excel stores when the user has nudged
 * the sheet-wide width.
 */
export function extractSheetFormatPr(xml: string): {
  defaultColWidthPx?: number;
  defaultRowHeightPx?: number;
} {
  const m = /<sheetFormatPr\b([^/>]*)\/?>/.exec(xml);
  if (!m) return {};
  const attrs = m[1] ?? "";
  const colMatch = /\bdefaultColWidth=("|')([^"']+)\1/.exec(attrs);
  const rowMatch = /\bdefaultRowHeight=("|')([^"']+)\1/.exec(attrs);
  const out: { defaultColWidthPx?: number; defaultRowHeightPx?: number } = {};
  if (colMatch) {
    const v = Number(colMatch[2]);
    if (Number.isFinite(v) && v > 0) out.defaultColWidthPx = charWidthToPx(v);
  }
  if (rowMatch) {
    const v = Number(rowMatch[2]);
    if (Number.isFinite(v) && v > 0) out.defaultRowHeightPx = ptToPx(v);
  }
  return out;
}

function parseFilterColumns(body: string): Map<number, FilterColumn> {
  const out = new Map<number, FilterColumn>();
  if (!body) return out;
  const fcRe = /<filterColumn\b([^>]*)(?:\/>|>([\s\S]*?)<\/filterColumn>)/g;
  let m: RegExpExecArray | null;
  while ((m = fcRe.exec(body)) !== null) {
    const attrs = m[1] ?? "";
    const inner = m[2] ?? "";
    const idMatch = /\bcolId=("|')([^"']+)\1/.exec(attrs);
    if (!idMatch) continue;
    const colId = Number(idMatch[2]);
    if (!Number.isInteger(colId) || colId < 0) continue;
    const fc = parseSingleFilterColumn(inner);
    if (fc) out.set(colId, fc);
  }
  return out;
}

function parseSingleFilterColumn(inner: string): FilterColumn | null {
  // <filters> with <filter val="..."/> children
  const filtersMatch = /<filters\b([^>]*)(?:\/>|>([\s\S]*?)<\/filters>)/.exec(inner);
  if (filtersMatch) {
    const attrs = filtersMatch[1] ?? "";
    const body = filtersMatch[2] ?? "";
    const blank = /\bblank=("|')(1|true)\1/.test(attrs);
    const values = new Set<string>();
    const fRe = /<filter\b[^/>]*\bval=("|')([^"']*)\1/g;
    let fm: RegExpExecArray | null;
    while ((fm = fRe.exec(body)) !== null) {
      values.add(fm[2] ?? "");
    }
    return { kind: "values", values, blank };
  }

  // <customFilters> with <customFilter operator="..." val="..."/>
  const cfMatch = /<customFilters\b([^>]*)(?:\/>|>([\s\S]*?)<\/customFilters>)/.exec(inner);
  if (cfMatch) {
    const attrs = cfMatch[1] ?? "";
    const body = cfMatch[2] ?? "";
    const combine: "and" | "or" = /\band=("|')(1|true)\1/.test(attrs) ? "and" : "or";
    const ops: CustomFilterOp[] = [];
    const opRe = /<customFilter\b([^/>]*)\/?>/g;
    let om: RegExpExecArray | null;
    while ((om = opRe.exec(body)) !== null) {
      const opAttrs = om[1] ?? "";
      const opMatch = /\boperator=("|')([^"']+)\1/.exec(opAttrs);
      const valMatch = /\bval=("|')([^"']*)\1/.exec(opAttrs);
      const operator = (opMatch?.[2] ?? "equal") as CustomFilterOp["operator"];
      const val = valMatch?.[2] ?? "";
      ops.push({ operator, val });
    }
    if (ops.length === 0) return null;
    return { kind: "custom", op1: ops[0]!, op2: ops[1], combine };
  }

  // <top10 top="1" percent="0" val="10" filterVal="..."/>
  const top10Match = /<top10\b([^/>]*)\/?>/.exec(inner);
  if (top10Match) {
    const attrs = top10Match[1] ?? "";
    const top = !/\btop=("|')0\1/.test(attrs); // default true
    const percent = /\bpercent=("|')(1|true)\1/.test(attrs);
    const valMatch = /\bval=("|')([^"']+)\1/.exec(attrs);
    const filterValMatch = /\bfilterVal=("|')([^"']+)\1/.exec(attrs);
    const n = valMatch ? Number(valMatch[2]) : 10;
    const filterVal = filterValMatch ? Number(filterValMatch[2]) : Number.NaN;
    return {
      kind: "top10",
      top,
      percent,
      n: Number.isFinite(n) ? n : 10,
      filterVal: Number.isFinite(filterVal) ? filterVal : 0,
    };
  }

  // <dynamicFilter type="..."/>
  const dynMatch = /<dynamicFilter\b([^/>]*)\/?>/.exec(inner);
  if (dynMatch) {
    const attrs = dynMatch[1] ?? "";
    const tMatch = /\btype=("|')([^"']+)\1/.exec(attrs);
    if (tMatch) {
      return { kind: "dynamic", type: tMatch[2] as DynamicFilterType };
    }
  }

  // <colorFilter dxfId="0" cellColor="1"/>
  // We persist the *resolved* color rather than the dxfId since the
  // dxf table is opaque in our model. Fall back to a dxfId stamp so a
  // round-trip stays semantically meaningful.
  const colorMatch = /<colorFilter\b([^/>]*)\/?>/.exec(inner);
  if (colorMatch) {
    const attrs = colorMatch[1] ?? "";
    const isCellColor = !/\bcellColor=("|')0\1/.test(attrs);
    const dxfMatch = /\bdxfId=("|')([^"']+)\1/.exec(attrs);
    return {
      kind: "color",
      argb: dxfMatch ? `dxf:${dxfMatch[2]}` : "FFFFFFFF",
      isCellColor,
    };
  }

  return null;
}

const ERROR_CODE_BY_NUM: Record<number, CellErrorCode> = {
  0x00: "#NULL!",
  0x07: "#DIV/0!",
  0x0f: "#VALUE!",
  0x17: "#REF!",
  0x1d: "#NAME?",
  0x24: "#NUM!",
  0x2a: "#N/A",
  0x2b: "#GETTING_DATA",
};

function mapErrorCode(v: unknown): CellValue {
  if (typeof v === "number" && ERROR_CODE_BY_NUM[v]) {
    return { kind: "error", code: ERROR_CODE_BY_NUM[v] };
  }
  if (typeof v === "string") {
    const code = v as CellErrorCode;
    return { kind: "error", code };
  }
  return { kind: "error", code: "#VALUE!" };
}

/**
 * Resolve a relationship target relative to the part that owns the
 * rels file. Container paths never have a leading slash; targets
 * starting with "/" are absolute (rare in workbook rels but legal).
 */
export function resolveTargetPath(ownerPartPath: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const ownerDir = ownerPartPath.includes("/") ? ownerPartPath.slice(0, ownerPartPath.lastIndexOf("/")) : "";
  const segments = (ownerDir ? ownerDir.split("/") : []).concat(target.split("/"));
  const stack: string[] = [];
  for (const seg of segments) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") {
      stack.pop();
    } else {
      stack.push(seg);
    }
  }
  return stack.join("/");
}
