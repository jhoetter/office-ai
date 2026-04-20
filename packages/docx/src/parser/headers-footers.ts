import { ooxml, type IdMinter } from "@officeai/core";
import type { BlockNode, HeaderFooterPart } from "../model/types.js";
import { DocxParseError } from "./errors.js";
import { attrOf, captureOpaque, elementEntries, findElementEntry, rootEntry } from "./xml-helpers.js";

/**
 * Parser for `word/header*.xml` / `word/footer*.xml` parts. The body shape is
 * identical to the main document body (paragraphs, tables, opaque blocks); we
 * delegate paragraph parsing to the same helper used by `parse.ts` so the
 * resulting model is uniform with the main body.
 *
 * Discovery: walk the relationships in `word/_rels/document.xml.rels` for the
 * two header / footer relationship types. For each hit, parse the target
 * part. The `target` (`default` / `first` / `even`) is read from the matching
 * `<w:headerReference w:type>` / `<w:footerReference w:type>` element inside
 * the document body's `<w:sectPr>`, falling back to `"default"` when no
 * matching reference is present (a defensive default — Word treats an
 * unspecified type as `default` too).
 */

const HEADER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const FOOTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
const MAIN_PART = "word/document.xml";

const HEADER_TAG = "w:hdr";
const FOOTER_TAG = "w:ftr";

const HEADER_REF_TAG = "w:headerReference";
const FOOTER_REF_TAG = "w:footerReference";

/**
 * Per-part discovered metadata: the relationship target (resolved to a part
 * path) and the section type the document's `<w:sectPr>` associates it with.
 */
interface DiscoveredRef {
  partPath: string;
  kind: "header" | "footer";
  target: HeaderFooterPart["target"];
}

/**
 * Discover every header/footer part referenced by the main document. The
 * caller is expected to have already loaded and validated the container; we
 * only consult relationships and the body's `<w:sectPr>` elements (we don't
 * re-parse the body).
 */
export function discoverHeaderFooterRefs(
  container: ooxml.OoxmlContainer,
  documentTree: unknown
): DiscoveredRef[] {
  const out: DiscoveredRef[] = [];
  if (!container.has(MAIN_PART)) return out;
  const rels = ooxml.RelationshipGraph.loadFor(container, MAIN_PART);

  // Build a map rId → "default" | "first" | "even" by scanning every
  // <w:sectPr> in the body. The body has at most a handful of section breaks
  // so this is cheap. If multiple references for the same rId disagree on
  // type (which would be malformed but tolerated by Word), we keep the first.
  const typeByRid = collectReferenceTypes(documentTree);

  // Walk header rels, then footer rels, in declaration order. Order matters
  // for byte-stability — we re-emit the parts in the same order on save.
  const seen = new Set<string>();
  for (const rel of rels.relationships) {
    if (rel.type !== HEADER_REL_TYPE && rel.type !== FOOTER_REL_TYPE) continue;
    const partPath = resolvePartPath(rel.target);
    if (seen.has(partPath)) continue;
    seen.add(partPath);
    if (!container.has(partPath)) continue;
    const recordedType = typeByRid.get(rel.id);
    out.push({
      partPath,
      kind: rel.type === HEADER_REL_TYPE ? "header" : "footer",
      target: recordedType ?? "default",
    });
  }
  return out;
}

/**
 * Parse the discovered header/footer parts into typed `HeaderFooterPart`
 * carriers. `parseParagraph` is injected (rather than imported) to avoid an
 * import cycle with `parse.ts`.
 */
export type ParseTableFn = (entry: Record<string, unknown>, mintNodeId: IdMinter) => BlockNode;

export function parseHeaderFooterParts(
  container: ooxml.OoxmlContainer,
  refs: ReadonlyArray<DiscoveredRef>,
  mintNodeId: IdMinter,
  parseParagraph: (entry: Record<string, unknown>, mintNodeId: IdMinter) => BlockNode,
  parseTable?: ParseTableFn
): HeaderFooterPart[] {
  const out: HeaderFooterPart[] = [];
  for (const ref of refs) {
    out.push(parseHeaderFooterPart(container, ref, mintNodeId, parseParagraph, parseTable));
  }
  return out;
}

function parseHeaderFooterPart(
  container: ooxml.OoxmlContainer,
  ref: DiscoveredRef,
  mintNodeId: IdMinter,
  parseParagraph: (entry: Record<string, unknown>, mintNodeId: IdMinter) => BlockNode,
  parseTable: ParseTableFn | undefined
): HeaderFooterPart {
  const xml = container.readText(ref.partPath);
  let tree: unknown;
  try {
    tree = ooxml.parseXml(xml);
  } catch (err) {
    throw new DocxParseError("invalid-xml", `Failed to parse ${ref.partPath}`, {
      partPath: ref.partPath,
      cause: err,
    });
  }
  const expectedTag = ref.kind === "header" ? HEADER_TAG : FOOTER_TAG;
  let rootEntryNode: Record<string, unknown>;
  try {
    rootEntryNode = rootEntry(tree, expectedTag);
  } catch (err) {
    throw new DocxParseError("missing-root", `Missing <${expectedTag}> root in ${ref.partPath}`, {
      partPath: ref.partPath,
      cause: err,
    });
  }

  const rootAttrs = readRootAttrs(rootEntryNode);
  const children = (rootEntryNode[expectedTag] as unknown[] | undefined) ?? [];
  const body: BlockNode[] = [];
  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    if (tag === "w:p") {
      body.push(parseParagraph(c, mintNodeId));
    } else if (tag === "w:tbl" && parseTable) {
      // Header/footer tables are now promoted to the same typed Table
      // model as body tables so the renderer + agent diff path stay
      // uniform. The default fallback to opaque-block remains as a
      // safety net when the typed parser isn't injected.
      body.push(parseTable(c, mintNodeId));
    } else {
      // SDT / opaque elements stay verbatim — typed SDT mutation
      // is out of scope. Unknown elements likewise become opaque
      // blocks per the body parser's defensive contract.
      body.push({ kind: "opaque-block", id: mintNodeId(), raw: captureOpaque(c) });
    }
  }

  return {
    kind: ref.kind,
    id: ref.partPath,
    partPath: ref.partPath,
    target: ref.target,
    rootAttrs,
    body,
  };
}

function readRootAttrs(entry: Record<string, unknown>): Record<string, string> {
  const attrs: Record<string, string> = {};
  const a = entry[":@"];
  if (!a || typeof a !== "object") return attrs;
  for (const [k, v] of Object.entries(a as Record<string, unknown>)) {
    const name = k.startsWith("@_") ? k.slice(2) : k;
    attrs[name] = String(v);
  }
  return attrs;
}

function collectReferenceTypes(documentTree: unknown): Map<string, HeaderFooterPart["target"]> {
  const out = new Map<string, HeaderFooterPart["target"]>();
  if (!Array.isArray(documentTree)) return out;
  const docEl = findElementEntry(documentTree as unknown[], "w:document");
  if (!docEl) return out;
  const body = findElementEntry((docEl["w:document"] as unknown[] | undefined) ?? [], "w:body");
  if (!body) return out;
  // sectPr can sit either at the body's top level or nested inside a
  // paragraph's pPr (the older "section break in paragraph" pattern). Walk
  // recursively from the body so we don't miss either form.
  walkAllSectPrs(body, out);
  return out;
}

function walkAllSectPrs(entry: Record<string, unknown>, out: Map<string, HeaderFooterPart["target"]>): void {
  const tag = ooxml.getTag(entry);
  if (tag === "w:sectPr") {
    walkSectPr(entry, out);
    return;
  }
  const children = entry[tag];
  if (!Array.isArray(children)) return;
  for (const c of elementEntries(children)) {
    walkAllSectPrs(c, out);
  }
}

function walkSectPr(sectPr: Record<string, unknown>, out: Map<string, HeaderFooterPart["target"]>): void {
  const children = (sectPr["w:sectPr"] as unknown[] | undefined) ?? [];
  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    if (tag !== HEADER_REF_TAG && tag !== FOOTER_REF_TAG) continue;
    const rid = attrOf(c, "r:id");
    if (!rid) continue;
    if (out.has(rid)) continue;
    const t = attrOf(c, "w:type");
    out.set(rid, normalizeTarget(t));
  }
}

function normalizeTarget(t: string | undefined): HeaderFooterPart["target"] {
  if (t === "first" || t === "even") return t;
  return "default";
}

function resolvePartPath(target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  return `word/${target}`;
}
