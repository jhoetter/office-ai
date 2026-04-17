import { defaultIdMinter, ooxml, sha256Hex, type IdMinter } from "@officeai/core";
import * as XLSX from "xlsx";
import {
  emptyDirty,
  type OpaquePart,
  type Sheet,
  type XlsxSnapshot,
  type XlsxWorkbook,
} from "../model/types.js";
import { XlsxParseError } from "./errors.js";

const WORKBOOK_PART = "xl/workbook.xml";
const WORKBOOK_RELS_PART = "xl/_rels/workbook.xml.rels";
const CONTENT_TYPES_PART = "[Content_Types].xml";

const DOC_REL_TYPES = {
  worksheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
  chartsheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
  dialogsheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
} as const;

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

const MODELED_PREFIXES = ["xl/worksheets/", "xl/_rels/", "xl/comments", "xl/threadedComments/"];

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

  const workbookRels = ooxml.RelationshipGraph.loadFor(container, WORKBOOK_PART);
  const sheets: Sheet[] = sheetEntries.map((entry, index) =>
    resolveSheet(entry, index, workbookRels, container, mintNodeId)
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

  const workbook: XlsxWorkbook = {
    id: mintNodeId(),
    sheets,
    partHashes,
    opaqueParts,
    date1904,
    workbookRootAttrs,
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

  return {
    id: mintNodeId(),
    sheetId: entry.sheetId,
    name: entry.name,
    index,
    state: entry.state,
    kind,
    partPath,
    ...(relsPartPath ? { relsPartPath } : {}),
  };
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
