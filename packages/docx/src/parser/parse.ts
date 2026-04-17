import { defaultIdMinter, ooxml, sha256Hex, type IdMinter, type NodeId } from "@officeai/core";
import type {
  BlockNode,
  CommentRangeEnd,
  CommentRangeStart,
  DocxComment,
  DocxDirtyFlags,
  DocxDocument,
  DocxSnapshot,
  HeaderFooterPart,
  Hyperlink,
  InlineNode,
  OpaqueBlock,
  OpaqueInline,
  Paragraph,
  ParagraphProperties,
  RevisionWrapper,
  Run,
  RunChild,
  RunProperties,
  SectionBreak,
} from "../model/types.js";
import { DocxParseError } from "./errors.js";
import { discoverHeaderFooterRefs, parseHeaderFooterParts } from "./headers-footers.js";
import { parseDrawing } from "./images.js";
import { parseMediaParts } from "./media.js";
import { parseRelationshipsParts } from "./relationships.js";
import { parseTable as parseTableTyped } from "./tables.js";
import {
  attrOf,
  captureOpaque,
  elementEntries,
  findElementEntry,
  readText,
  rootEntry,
} from "./xml-helpers.js";

export interface ParseOptions {
  readonly idMinter?: IdMinter;
}

const MAIN_PART = "word/document.xml";
const COMMENTS_PART = "word/comments.xml";
const COMMENTS_EXTENDED_PART = "word/commentsExtended.xml";

const DOC_ROOT_ATTR_KEYS = [
  "xmlns:w",
  "xmlns:r",
  "xmlns:m",
  "xmlns:v",
  "xmlns:wp",
  "xmlns:wp14",
  "xmlns:w14",
  "xmlns:w15",
  "xmlns:w16",
  "xmlns:w16cex",
  "xmlns:w16cid",
  "xmlns:w16se",
  "xmlns:mc",
  "xmlns:o",
  "xmlns:wpc",
  "xmlns:wpg",
  "xmlns:wpi",
  "xmlns:wne",
  "xmlns:wps",
  "xmlns:cx",
  "xmlns:cx1",
  "xmlns:cx2",
  "xmlns:cx3",
  "xmlns:cx4",
  "xmlns:cx5",
  "xmlns:cx6",
  "xmlns:cx7",
  "xmlns:cx8",
  "mc:Ignorable",
  "xml:space",
];

export async function parseDocx(
  input: ArrayBuffer | Uint8Array,
  opts: ParseOptions = {}
): Promise<DocxSnapshot> {
  let container: ooxml.OoxmlContainer;
  try {
    container = await ooxml.OoxmlContainer.load(input);
  } catch (err) {
    throw new DocxParseError("invalid-zip", "Failed to read DOCX as a zip archive", { cause: err });
  }
  if (!container.has(MAIN_PART)) {
    throw new DocxParseError("missing-main-part", `Missing required part: ${MAIN_PART}`, {
      partPath: MAIN_PART,
    });
  }
  const mintNodeId: IdMinter = opts.idMinter ?? defaultIdMinter;

  const documentXml = container.readText(MAIN_PART);
  let documentTree: unknown;
  try {
    documentTree = ooxml.parseXml(documentXml);
  } catch (err) {
    throw new DocxParseError("invalid-xml", `Failed to parse ${MAIN_PART}`, {
      partPath: MAIN_PART,
      cause: err,
    });
  }

  let docEntry: Record<string, unknown>;
  try {
    docEntry = rootEntry(documentTree, "w:document");
  } catch (err) {
    throw new DocxParseError("missing-root", "Missing <w:document> root", {
      partPath: MAIN_PART,
      cause: err,
    });
  }

  const documentRootAttrs = readDocRootAttrs(docEntry);
  const bodyEntry = findElementEntry(docEntry["w:document"] as unknown[], "w:body");
  if (!bodyEntry) {
    throw new DocxParseError("missing-body", "Missing <w:body> in document", { partPath: MAIN_PART });
  }

  const body = parseBody(bodyEntry, mintNodeId);
  const baseComments = parseComments(container, mintNodeId);
  const extended = parseCommentsExtended(container);
  const comments = applyCommentsExtended(baseComments, extended);
  const headerFooterRefs = discoverHeaderFooterRefs(container, documentTree);
  const headersAndFooters: HeaderFooterPart[] = parseHeaderFooterParts(
    container,
    headerFooterRefs,
    mintNodeId,
    parseParagraph
  );

  const media = parseMediaParts(container);
  const relationships = parseRelationshipsParts(container);

  const root: DocxDocument = {
    id: mintNodeId(),
    body,
    comments,
    headersAndFooters,
    media,
    relationships,
    documentRootAttrs,
  };

  const partHashes: Record<string, string> = {};
  for (const path of container.parts.keys()) {
    partHashes[path] = sha256Hex(container.readBytes(path));
  }

  const dirty: DocxDirtyFlags = {
    body: false,
    comments: false,
    rels: false,
    contentTypes: false,
    commentsExtended: false,
    headersAndFooters: new Set<string>(),
    media: new Set<string>(),
    relationships: new Set<string>(),
  };

  return {
    format: "docx",
    revision: 0,
    root,
    partHashes,
    container,
    dirty,
  };
}

function readDocRootAttrs(docEntry: Record<string, unknown>): Record<string, string> {
  const attrs: Record<string, string> = {};
  const a = docEntry[":@"];
  if (!a || typeof a !== "object") return attrs;
  const map = a as Record<string, unknown>;
  for (const [k, v] of Object.entries(map)) {
    const name = k.startsWith("@_") ? k.slice(2) : k;
    if (DOC_ROOT_ATTR_KEYS.indexOf(name) !== -1 || name.startsWith("xmlns:")) {
      attrs[name] = String(v);
    } else {
      attrs[name] = String(v);
    }
  }
  return attrs;
}

function parseBody(bodyEntry: Record<string, unknown>, mintNodeId: IdMinter): BlockNode[] {
  const out: BlockNode[] = [];
  const children = (bodyEntry["w:body"] as unknown[] | undefined) ?? [];
  for (const entry of elementEntries(children)) {
    const tag = ooxml.getTag(entry);
    switch (tag) {
      case "w:p":
        out.push(parseParagraph(entry, mintNodeId));
        break;
      case "w:tbl":
        out.push(parseTableTyped(entry, mintNodeId, parseParagraph));
        break;
      case "w:sectPr":
        out.push(parseSectionBreak(entry, mintNodeId));
        break;
      default:
        out.push(parseOpaqueBlock(entry, mintNodeId));
        break;
    }
  }
  return out;
}

function parseSectionBreak(entry: Record<string, unknown>, mintNodeId: IdMinter): SectionBreak {
  return { kind: "section-break", id: mintNodeId(), raw: captureOpaque(entry) };
}

function parseOpaqueBlock(entry: Record<string, unknown>, mintNodeId: IdMinter): OpaqueBlock {
  return { kind: "opaque-block", id: mintNodeId(), raw: captureOpaque(entry) };
}

function parseParagraph(entry: Record<string, unknown>, mintNodeId: IdMinter): Paragraph {
  const children = (entry["w:p"] as unknown[] | undefined) ?? [];
  const pPr = findElementEntry(children, "w:pPr");
  const properties = pPr ? parseParagraphProperties(pPr) : {};
  const inlines: InlineNode[] = [];
  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    if (tag === "w:pPr") continue;
    inlines.push(parseInline(c, mintNodeId));
  }
  return {
    kind: "paragraph",
    id: mintNodeId(),
    properties,
    children: inlines,
  };
}

function parseParagraphProperties(entry: Record<string, unknown>): ParagraphProperties {
  const children = (entry["w:pPr"] as unknown[] | undefined) ?? [];
  const props: {
    -readonly [K in keyof ParagraphProperties]: ParagraphProperties[K];
  } = {};
  const opaqueProps: ReturnType<typeof captureOpaque>[] = [];

  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    switch (tag) {
      case "w:pStyle": {
        const v = attrOf(c, "w:val");
        if (v) props.styleId = v;
        break;
      }
      case "w:jc": {
        const v = attrOf(c, "w:val");
        const map: Record<string, NonNullable<ParagraphProperties["alignment"]>> = {
          left: "left",
          start: "left",
          right: "right",
          end: "right",
          center: "center",
          both: "justify",
          distribute: "justify",
        };
        if (v && map[v]) props.alignment = map[v];
        else opaqueProps.push(captureOpaque(c));
        break;
      }
      case "w:ind": {
        const ind: NonNullable<ParagraphProperties["indentation"]> = {};
        const left = attrOf(c, "w:left") ?? attrOf(c, "w:start");
        const right = attrOf(c, "w:right") ?? attrOf(c, "w:end");
        const firstLine = attrOf(c, "w:firstLine");
        const hanging = attrOf(c, "w:hanging");
        if (left !== undefined) (ind as Record<string, number>).left = Number(left);
        if (right !== undefined) (ind as Record<string, number>).right = Number(right);
        if (firstLine !== undefined) (ind as Record<string, number>).firstLine = Number(firstLine);
        if (hanging !== undefined) (ind as Record<string, number>).hanging = Number(hanging);
        if (Object.keys(ind).length > 0) props.indentation = ind;
        break;
      }
      case "w:spacing": {
        const sp: NonNullable<ParagraphProperties["spacing"]> = {};
        const before = attrOf(c, "w:before");
        const after = attrOf(c, "w:after");
        const line = attrOf(c, "w:line");
        const lineRule = attrOf(c, "w:lineRule");
        if (before !== undefined) (sp as Record<string, number>).before = Number(before);
        if (after !== undefined) (sp as Record<string, number>).after = Number(after);
        if (line !== undefined) (sp as Record<string, number>).line = Number(line);
        if (lineRule === "auto" || lineRule === "exact" || lineRule === "atLeast") {
          (sp as { lineRule?: typeof lineRule }).lineRule = lineRule;
        }
        if (Object.keys(sp).length > 0) props.spacing = sp;
        break;
      }
      case "w:numPr": {
        const numIdEl = findElementEntry((c["w:numPr"] as unknown[] | undefined) ?? [], "w:numId");
        const ilvlEl = findElementEntry((c["w:numPr"] as unknown[] | undefined) ?? [], "w:ilvl");
        const numId = numIdEl ? Number(attrOf(numIdEl, "w:val") ?? "0") : 0;
        const ilvl = ilvlEl ? Number(attrOf(ilvlEl, "w:val") ?? "0") : 0;
        props.numbering = { numId, ilvl };
        opaqueProps.push(captureOpaque(c));
        break;
      }
      default:
        opaqueProps.push(captureOpaque(c));
        break;
    }
  }

  if (opaqueProps.length > 0) props.opaqueProps = opaqueProps;
  return props;
}

function parseInline(entry: Record<string, unknown>, mintNodeId: IdMinter): InlineNode {
  const tag = ooxml.getTag(entry);
  switch (tag) {
    case "w:r":
      return parseRun(entry, mintNodeId);
    case "w:hyperlink":
      return parseHyperlink(entry, mintNodeId);
    case "w:ins":
    case "w:del":
      return parseRevisionWrapper(entry, mintNodeId);
    case "w:commentRangeStart":
      return {
        kind: "comment-range-start",
        id: mintNodeId(),
        commentId: attrOf(entry, "w:id") ?? "",
      } satisfies CommentRangeStart;
    case "w:commentRangeEnd":
      return {
        kind: "comment-range-end",
        id: mintNodeId(),
        commentId: attrOf(entry, "w:id") ?? "",
      } satisfies CommentRangeEnd;
    default:
      return {
        kind: "opaque-inline",
        id: mintNodeId(),
        raw: captureOpaque(entry),
      } satisfies OpaqueInline;
  }
}

function parseRun(entry: Record<string, unknown>, mintNodeId: IdMinter): Run {
  const children = (entry["w:r"] as unknown[] | undefined) ?? [];
  const rPr = findElementEntry(children, "w:rPr");
  const properties = rPr ? parseRunProperties(rPr) : {};
  const runChildren: RunChild[] = [];
  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    if (tag === "w:rPr") continue;
    runChildren.push(parseRunChild(c, mintNodeId));
  }
  return { kind: "run", id: mintNodeId(), properties, children: runChildren };
}

function parseRunProperties(entry: Record<string, unknown>): RunProperties {
  const children = (entry["w:rPr"] as unknown[] | undefined) ?? [];
  const props: { -readonly [K in keyof RunProperties]: RunProperties[K] } = {};
  const opaqueProps: ReturnType<typeof captureOpaque>[] = [];
  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    switch (tag) {
      case "w:b":
        props.bold = boolAttr(attrOf(c, "w:val"), true);
        break;
      case "w:i":
        props.italic = boolAttr(attrOf(c, "w:val"), true);
        break;
      case "w:u": {
        const v = attrOf(c, "w:val");
        if (!v || v === "none") props.underline = false;
        else if (v === "single") props.underline = true;
        else props.underline = v;
        break;
      }
      case "w:strike":
        props.strike = boolAttr(attrOf(c, "w:val"), true);
        break;
      case "w:rFonts": {
        const ascii = attrOf(c, "w:ascii");
        if (ascii) props.fontFamily = ascii;
        opaqueProps.push(captureOpaque(c));
        break;
      }
      case "w:sz": {
        const v = attrOf(c, "w:val");
        if (v) props.fontSize = Number(v);
        break;
      }
      case "w:color": {
        const v = attrOf(c, "w:val");
        if (v) props.color = v;
        break;
      }
      case "w:highlight": {
        const v = attrOf(c, "w:val");
        if (v) props.highlight = v;
        break;
      }
      default:
        opaqueProps.push(captureOpaque(c));
    }
  }
  if (opaqueProps.length > 0) props.opaqueProps = opaqueProps;
  return props;
}

function parseRunChild(entry: Record<string, unknown>, mintNodeId: IdMinter): RunChild {
  const tag = ooxml.getTag(entry);
  switch (tag) {
    case "w:t":
    case "w:delText": {
      const text = readText(entry);
      const xmlSpacePreserve = attrOf(entry, "xml:space") === "preserve";
      return {
        kind: "text",
        id: mintNodeId(),
        text,
        xmlSpacePreserve,
        ...(tag === "w:delText" ? { isDelText: true } : {}),
      };
    }
    case "w:br": {
      const t = attrOf(entry, "w:type");
      const breakType = t === "page" || t === "column" || t === "textWrapping" ? t : undefined;
      return { kind: "break", id: mintNodeId(), ...(breakType ? { breakType } : {}) };
    }
    case "w:tab":
      return { kind: "tab", id: mintNodeId() };
    case "w:drawing":
      return parseDrawing(entry, mintNodeId);
    default:
      return { kind: "opaque", id: mintNodeId(), raw: captureOpaque(entry) };
  }
}

function parseHyperlink(entry: Record<string, unknown>, mintNodeId: IdMinter): Hyperlink {
  const rId = attrOf(entry, "r:id");
  const anchor = attrOf(entry, "w:anchor");
  const children = (entry["w:hyperlink"] as unknown[] | undefined) ?? [];
  const runs: Run[] = [];
  for (const c of elementEntries(children)) {
    if (ooxml.getTag(c) === "w:r") runs.push(parseRun(c, mintNodeId));
  }
  return {
    kind: "hyperlink",
    id: mintNodeId(),
    ...(rId ? { relationshipId: rId } : {}),
    ...(anchor ? { anchor } : {}),
    children: runs,
  };
}

function parseRevisionWrapper(entry: Record<string, unknown>, mintNodeId: IdMinter): RevisionWrapper {
  const tag = ooxml.getTag(entry);
  const childrenSrc = (entry[tag] as unknown[] | undefined) ?? [];
  const inlines: InlineNode[] = [];
  for (const c of elementEntries(childrenSrc)) {
    inlines.push(parseInline(c, mintNodeId));
  }
  return {
    kind: "revision",
    id: mintNodeId(),
    revisionType: tag === "w:ins" ? "ins" : "del",
    author: attrOf(entry, "w:author") ?? "",
    date: attrOf(entry, "w:date") ?? "",
    revisionId: attrOf(entry, "w:id") ?? "",
    children: inlines,
  };
}

function boolAttr(value: string | undefined, fallback: boolean): boolean {
  if (value === undefined) return fallback;
  if (value === "false" || value === "0" || value === "off") return false;
  return true;
}

function parseComments(container: ooxml.OoxmlContainer, mintNodeId: IdMinter): DocxComment[] {
  if (!container.has(COMMENTS_PART)) return [];
  let tree: unknown;
  try {
    tree = ooxml.parseXml(container.readText(COMMENTS_PART));
  } catch (err) {
    throw new DocxParseError("invalid-xml", "Failed to parse comments.xml", {
      partPath: COMMENTS_PART,
      cause: err,
    });
  }
  if (!Array.isArray(tree)) return [];
  const root = findElementEntry(tree as unknown[], "w:comments");
  if (!root) return [];
  const out: DocxComment[] = [];
  for (const c of elementEntries((root["w:comments"] as unknown[] | undefined) ?? [])) {
    if (ooxml.getTag(c) !== "w:comment") continue;
    const id = attrOf(c, "w:id") ?? "";
    const author = attrOf(c, "w:author") ?? "";
    const initials = attrOf(c, "w:initials");
    const date = attrOf(c, "w:date") ?? "";
    const body: BlockNode[] = [];
    let paraId: string | undefined;
    for (const p of elementEntries((c["w:comment"] as unknown[] | undefined) ?? [])) {
      if (ooxml.getTag(p) !== "w:p") continue;
      if (paraId === undefined) {
        const pid = attrOf(p, "w14:paraId");
        if (pid) paraId = pid;
      }
      body.push(parseParagraph(p, mintNodeId));
    }
    out.push({
      id,
      author,
      ...(initials ? { initials } : {}),
      date,
      body,
      ...(paraId ? { paraId } : {}),
    });
  }
  return out;
}

interface ExtendedRecord {
  paraId: string;
  done?: boolean;
  parentParaId?: string;
}

/**
 * Parse `word/commentsExtended.xml` if present. Returns a map keyed by
 * `w15:paraId` of the comment's first body paragraph. Missing part →
 * empty map; unparseable part → empty map (treated as "no extended
 * metadata", not a hard error, since older Word docs and many third-party
 * tools omit it).
 */
function parseCommentsExtended(container: ooxml.OoxmlContainer): Map<string, ExtendedRecord> {
  const out = new Map<string, ExtendedRecord>();
  if (!container.has(COMMENTS_EXTENDED_PART)) return out;
  let tree: unknown;
  try {
    tree = ooxml.parseXml(container.readText(COMMENTS_EXTENDED_PART));
  } catch {
    return out;
  }
  if (!Array.isArray(tree)) return out;
  const root = findElementEntry(tree as unknown[], "w15:commentsEx");
  if (!root) return out;
  for (const c of elementEntries((root["w15:commentsEx"] as unknown[] | undefined) ?? [])) {
    if (ooxml.getTag(c) !== "w15:commentEx") continue;
    const paraId = attrOf(c, "w15:paraId");
    if (!paraId) continue;
    const doneStr = attrOf(c, "w15:done");
    const parent = attrOf(c, "w15:parentPaIdRef");
    const rec: ExtendedRecord = { paraId };
    if (doneStr === "1" || doneStr === "true") rec.done = true;
    if (parent) rec.parentParaId = parent;
    out.set(paraId, rec);
  }
  return out;
}

/**
 * Project `commentsExtended` records onto the parsed comment list. Comments
 * with no matching paraId pass through unchanged.
 */
function applyCommentsExtended(
  comments: DocxComment[],
  extended: Map<string, ExtendedRecord>
): DocxComment[] {
  if (extended.size === 0) return comments;
  const paraIdToCommentId = new Map<string, string>();
  for (const c of comments) {
    if (c.paraId) paraIdToCommentId.set(c.paraId, c.id);
  }
  return comments.map((c) => {
    if (!c.paraId) return c;
    const ext = extended.get(c.paraId);
    if (!ext) return c;
    const next: DocxComment = { ...c };
    if (ext.done === true) (next as { resolved?: boolean }).resolved = true;
    if (ext.parentParaId) {
      const parentCommentId = paraIdToCommentId.get(ext.parentParaId);
      if (parentCommentId) (next as { parentId?: string }).parentId = parentCommentId;
    }
    return next;
  });
}

export type { NodeId };
