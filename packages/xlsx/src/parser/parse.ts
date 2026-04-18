import { defaultIdMinter, ooxml, sha256Hex, type IdMinter } from "@officeai/core";
import * as XLSX from "xlsx";
import { cellKey, formatA1 } from "../model/refs.js";
import { defaultStyleTable, type StyleTable } from "../model/style-table.js";
import {
  emptyDirty,
  type Cell,
  type CellErrorCode,
  type CellValue,
  type Comment,
  type Formula,
  type MergedCell,
  type OpaquePart,
  type Sheet,
  type XlsxSnapshot,
  type XlsxWorkbook,
} from "../model/types.js";
import { parseCommentsPart } from "./comments.js";
import { XlsxParseError } from "./errors.js";
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

  const { workbookRootAttrs, date1904, sheetEntries } = parseWorkbookXml(container);

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
  const sheets: Sheet[] = sheetEntries.map((entry, index) =>
    resolveSheet(entry, index, workbookRels, container, sheetjsBook, mintNodeId)
  );

  const contentTypes = ooxml.ContentTypes.load(container);
  const ctMap = new Map<string, string>();
  for (const o of contentTypes.overrides) {
    const partName = o.partName.startsWith("/") ? o.partName.slice(1) : o.partName;
    ctMap.set(partName, o.contentType);
  }

  const opaqueParts = new Map<string, OpaquePart>();
  const partHashes: Record<string, string> = {};
  for (const [path, part] of container.parts) {
    const hash = sha256Hex(part.bytes);
    partHashes[path] = hash;
    if (isModeledPath(path)) continue;
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

  const workbook: XlsxWorkbook = {
    id: mintNodeId(),
    sheets,
    partHashes,
    opaqueParts,
    date1904,
    workbookRootAttrs,
    styles,
    sheetjs: sheetjsBook,
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

function parseWorkbookXml(container: ooxml.OoxmlContainer): {
  workbookRootAttrs: Record<string, string>;
  date1904: boolean;
  sheetEntries: RawSheetEntry[];
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

  return { workbookRootAttrs, date1904, sheetEntries };
}

function resolveSheet(
  entry: RawSheetEntry,
  index: number,
  workbookRels: ooxml.RelationshipGraph,
  container: ooxml.OoxmlContainer,
  sheetjsBook: XLSX.WorkBook,
  mintNodeId: IdMinter
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
  const styleIdByRef =
    kind === "worksheet" && container.has(partPath)
      ? extractStyleIdsFromXml(container.readText(partPath))
      : undefined;
  const { cells, merges } =
    kind === "worksheet" && ws ? extractCellsAndMerges(ws, styleIdByRef) : EMPTY_CELL_DATA;

  const { commentsPartPath, comments, commentAuthors } = resolveComments(container, partPath);

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
    ...(commentsPartPath ? { commentsPartPath } : {}),
  };
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
