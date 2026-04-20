import { defaultIdMinter, ooxml, sha256Hex, type IdMinter, type NodeId } from "@officeai/core";
import type {
  BlockNode,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  DocxComment,
  DocxDirtyFlags,
  DocxDocument,
  DocxSnapshot,
  EmbeddedSpreadsheet,
  HeaderFooterPart,
  Hyperlink,
  InlineNode,
  OpaqueBlock,
  OpaqueInline,
  PageNumberFieldLeaf,
  Paragraph,
  ParagraphProperties,
  RevisionWrapper,
  Run,
  RunChild,
  RunProperties,
  SectionBreak,
  WrapperMarker,
} from "../model/types.js";
import { DocxParseError } from "./errors.js";
import { discoverHeaderFooterRefs, parseHeaderFooterParts } from "./headers-footers.js";
import { parseChartParts } from "./charts.js";
import { parseDrawing } from "./images.js";
import { parseEmbeddingParts, parseMediaParts } from "./media.js";
import { parseNumberingPart } from "./numbering.js";
import { parseRelationshipsParts } from "./relationships.js";
import { parseSectionProperties } from "./sections.js";
import { parseStylesPart } from "./styles.js";
import { parseThemePart } from "./theme.js";
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

/**
 * Resolver passed to `parseDrawing` so chart drawings get promoted
 * from `OpaqueDrawing` to the typed `ChartDrawing` shape. Set by
 * `parseDocx` once relationships have been parsed; cleared at the
 * end so re-entrant parses don't bleed state. We thread it through
 * a module-private slot rather than every parse signature because
 * the resolver is needed only inside the run-child parser, several
 * call levels deep.
 */
let currentChartResolver: ((relId: string) => string | undefined) | undefined;
/**
 * Generic relationship resolver scoped to the active main-document
 * parse. Lets `parseRunChild` look up arbitrary `r:id` targets (OLE
 * embeds, image preview blips, …) without threading the relationship
 * list through every intermediate function. Mirrors the chart resolver
 * pattern; cleared in the same `finally` block.
 */
let currentDocRelResolver: ((relId: string) => string | undefined) | undefined;

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

  // Parse relationships + chart parts FIRST so the body parser can
  // promote `<w:drawing>` chart references to typed `ChartDrawing`
  // leaves via the module-private resolver.
  const relationships = parseRelationshipsParts(container);
  const charts = parseChartParts(container, relationships, mintNodeId);
  const docRels = relationships.get("word/document.xml") ?? [];
  const chartByRelId = new Map<string, string>();
  for (const r of docRels) {
    if (r.type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart") {
      chartByRelId.set(r.id, resolveDocPath(r.target));
    }
  }
  currentChartResolver = (relId: string) => chartByRelId.get(relId);
  const docRelTargets = new Map<string, string>();
  for (const r of docRels) docRelTargets.set(r.id, resolveDocPath(r.target));
  currentDocRelResolver = (relId: string) => docRelTargets.get(relId);

  let body: BlockNode[];
  let headersAndFooters: HeaderFooterPart[];
  try {
    body = parseBody(bodyEntry, mintNodeId);
    const headerFooterRefs = discoverHeaderFooterRefs(container, documentTree);
    headersAndFooters = parseHeaderFooterParts(
      container,
      headerFooterRefs,
      mintNodeId,
      parseParagraph,
      (entry, mint) => parseTableTyped(entry, mint, parseParagraph)
    );
  } finally {
    currentChartResolver = undefined;
    currentDocRelResolver = undefined;
  }

  const baseComments = parseComments(container, mintNodeId);
  const extended = parseCommentsExtended(container);
  const comments = applyCommentsExtended(baseComments, extended);

  const media = parseMediaParts(container);
  const embeddings = parseEmbeddingParts(container);
  const numbering = parseNumberingPart(container);
  const styles = parseStylesPart(container);
  const theme = parseThemePart(container);

  const root: DocxDocument = {
    id: mintNodeId(),
    body,
    comments,
    headersAndFooters,
    media,
    relationships,
    charts,
    embeddings,
    ...(numbering ? { numbering } : {}),
    ...(styles ? { styles } : {}),
    ...(theme ? { theme } : {}),
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
    numbering: false,
    styles: false,
    charts: new Set<string>(),
    embeddings: new Set<string>(),
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

/**
 * Resolve a relationship `target` declared in `word/_rels/document.xml.rels`
 * (which is relative to `word/`) into a full part path under the OOXML
 * package, with no leading slash.
 */
function resolveDocPath(target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const segments = ("word/" + target).split("/");
  const out: string[] = [];
  for (const seg of segments) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") {
      out.pop();
      continue;
    }
    out.push(seg);
  }
  return out.join("/");
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
        appendOpaqueOrLifted(entry, mintNodeId, out);
        break;
    }
  }
  return out;
}

/**
 * Body-level entry point for non-paragraph / non-table / non-sectPr
 * elements (`<w:sdt>`, `mc:AlternateContent`, `<w:fldSimple>`,
 * `<w:smartTag>`, `<w:customXml>`).
 *
 * For carriers we know how to crack open (`blockContentSlot` returns
 * a non-null slot), AND that crack open to typed body blocks
 * (paragraphs, tables, section breaks), we LIFT the inner blocks into
 * `body` directly, bracketed by `wrapper-marker` blocks so the
 * serializer can rebuild the carrier's envelope on a body-dirty
 * round-trip.
 *
 * This is what fixes the "TOC sits on its own page even though the
 * heading is short" symptom: the SDT used to be one giant atom, now
 * it's a sequence of regular body blocks the page chunker can flow
 * across pages.
 */
function appendOpaqueOrLifted(entry: Record<string, unknown>, mintNodeId: IdMinter, out: BlockNode[]): void {
  const lifted = liftOpaqueBlock(entry, mintNodeId);
  if (lifted) {
    for (const block of lifted) out.push(block);
    return;
  }
  out.push(parseOpaqueBlock(entry, mintNodeId));
}

let WRAPPER_ID_COUNTER = 0;
function nextWrapperId(): string {
  WRAPPER_ID_COUNTER += 1;
  return `wrap-${WRAPPER_ID_COUNTER}`;
}

/**
 * Try to lift a body-level content-wrapper carrier. Returns `null`
 * when the entry is not a known wrapper or its content slot does not
 * contain typed body blocks (in which case the caller falls back to
 * the legacy `parseOpaqueBlock` path that produces a single atom).
 *
 * Returns a flat list `[ wrapper-begin, ...inner blocks, wrapper-end ]`
 * when the lift succeeds.
 */
function liftOpaqueBlock(entry: Record<string, unknown>, mintNodeId: IdMinter): BlockNode[] | null {
  const slot = blockContentSlot(entry);
  if (slot === null || slot.length === 0) return null;
  // Only lift when the slot contains at least one typed body block.
  // Pure-metadata slots (e.g. an `<w:sdt>` whose content is just a
  // single `<w:r>` with a field result) are still better rendered as
  // a chip rather than re-flowed into body.
  let hasTypedBlock = false;
  for (const child of slot) {
    const childTag = ooxml.getTag(child);
    if (childTag === "w:p" || childTag === "w:tbl" || childTag === "w:sectPr") {
      hasTypedBlock = true;
      break;
    }
  }
  if (!hasTypedBlock) return null;

  const wrapperRaw = captureOpaque(entry);
  const wrapperId = nextWrapperId();

  const out: BlockNode[] = [];
  const begin: WrapperMarker = {
    kind: "wrapper-marker",
    id: mintNodeId(),
    side: "begin",
    wrapperId,
    wrapperRaw,
  };
  out.push(begin);
  for (const child of slot) {
    const childTag = ooxml.getTag(child);
    switch (childTag) {
      case "w:p":
        out.push(parseParagraph(child, mintNodeId));
        break;
      case "w:tbl":
        out.push(parseTableTyped(child, mintNodeId, parseParagraph));
        break;
      case "w:sectPr":
        out.push(parseSectionBreak(child, mintNodeId));
        break;
      default: {
        // Nested wrapper — recurse so SDTs inside SDTs also lift.
        const nested = liftOpaqueBlock(child, mintNodeId);
        if (nested) {
          for (const n of nested) out.push(n);
        } else {
          out.push(parseOpaqueBlock(child, mintNodeId));
        }
        break;
      }
    }
  }
  const end: WrapperMarker = {
    kind: "wrapper-marker",
    id: mintNodeId(),
    side: "end",
    wrapperId,
    wrapperRaw,
  };
  out.push(end);
  return out;
}

function parseSectionBreak(entry: Record<string, unknown>, mintNodeId: IdMinter): SectionBreak {
  const properties = parseSectionProperties(entry);
  return {
    kind: "section-break",
    id: mintNodeId(),
    properties,
    raw: captureOpaque(entry),
  };
}

/**
 * Tags whose inner content (or the content of a designated child slot) is
 * regular block-level OOXML and can be parsed as typed `BlockNode`s. The
 * wrapper itself is still preserved verbatim through `OpaqueBlock.raw`; we
 * only attach the typed projection so the renderer can show the wrapped
 * paragraphs as real headings/paragraphs instead of an opaque chip.
 *
 * Returns `null` when the entry is not an unwrappable carrier (so the
 * caller falls back to the legacy "no children" path).
 */
function blockContentSlot(entry: Record<string, unknown>): Array<Record<string, unknown>> | null {
  const tag = ooxml.getTag(entry);
  const children = (entry[tag] as unknown[] | undefined) ?? [];
  switch (tag) {
    case "w:sdt": {
      const content = findElementEntry(children, "w:sdtContent");
      if (!content) return null;
      return elementEntries((content["w:sdtContent"] as unknown[] | undefined) ?? []);
    }
    case "mc:AlternateContent": {
      const choice = findElementEntry(children, "mc:Choice") ?? findElementEntry(children, "mc:Fallback");
      if (!choice) return null;
      const choiceTag = ooxml.getTag(choice);
      return elementEntries((choice[choiceTag] as unknown[] | undefined) ?? []);
    }
    case "w:sdtContent":
    case "mc:Choice":
    case "mc:Fallback":
    case "w:fldSimple":
    case "w:smartTag":
    case "w:customXml":
      return elementEntries(children);
    default:
      return null;
  }
}

function parseOpaqueBlock(entry: Record<string, unknown>, mintNodeId: IdMinter): OpaqueBlock {
  const raw = captureOpaque(entry);
  const slot = blockContentSlot(entry);
  if (slot === null || slot.length === 0) {
    return { kind: "opaque-block", id: mintNodeId(), raw };
  }
  const children: BlockNode[] = [];
  for (const child of slot) {
    const childTag = ooxml.getTag(child);
    switch (childTag) {
      case "w:p":
        children.push(parseParagraph(child, mintNodeId));
        break;
      case "w:tbl":
        children.push(parseTableTyped(child, mintNodeId, parseParagraph));
        break;
      case "w:sectPr":
        children.push(parseSectionBreak(child, mintNodeId));
        break;
      default:
        children.push(parseOpaqueBlock(child, mintNodeId));
        break;
    }
  }
  if (children.length === 0) {
    return { kind: "opaque-block", id: mintNodeId(), raw };
  }
  return { kind: "opaque-block", id: mintNodeId(), raw, children };
}

function parseOpaqueInline(entry: Record<string, unknown>, mintNodeId: IdMinter): OpaqueInline {
  const raw = captureOpaque(entry);
  const slot = blockContentSlot(entry);
  if (slot === null || slot.length === 0) {
    return { kind: "opaque-inline", id: mintNodeId(), raw };
  }
  const children: InlineNode[] = [];
  for (const child of slot) {
    children.push(parseInline(child, mintNodeId));
  }
  if (children.length === 0) {
    return { kind: "opaque-inline", id: mintNodeId(), raw };
  }
  return { kind: "opaque-inline", id: mintNodeId(), raw, children };
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

/**
 * Parse a `<w:pPr>` entry into typed paragraph properties. Exported
 * (alongside `parseRunProperties`) so the styles parser can reuse the
 * same OOXML → typed shape — the cascade resolver expects identical
 * shapes for `docDefaults.pPrDefault`, `style.pPr`, and `paragraph.pPr`.
 */
export function parseParagraphProperties(entry: Record<string, unknown>): ParagraphProperties {
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
      case "w:keepNext": {
        // Pagination flag: stay on same page as next block. The
        // element is a typed projection only — we keep the original
        // entry in opaqueProps so the serializer's existing emit path
        // re-emits it byte-identical.
        props.keepNext = ooxmlBoolFlag(c);
        opaqueProps.push(captureOpaque(c));
        break;
      }
      case "w:keepLines": {
        props.keepLines = ooxmlBoolFlag(c);
        opaqueProps.push(captureOpaque(c));
        break;
      }
      case "w:pageBreakBefore": {
        props.pageBreakBefore = ooxmlBoolFlag(c);
        opaqueProps.push(captureOpaque(c));
        break;
      }
      case "w:widowControl": {
        // Note: widow control defaults to TRUE in Word; the typical
        // emit pattern is `<w:widowControl w:val="0"/>` to disable.
        props.widowControl = ooxmlBoolFlag(c);
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
    case "w:r": {
      const promoted = tryParseRunAsCommentReference(entry, mintNodeId);
      if (promoted) return promoted;
      return parseRun(entry, mintNodeId);
    }
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
    case "w:fldSimple": {
      const promoted = tryParseFldSimpleAsPageNumber(entry, mintNodeId);
      if (promoted) return promoted;
      return parseOpaqueInline(entry, mintNodeId);
    }
    default:
      return parseOpaqueInline(entry, mintNodeId);
  }
}

/**
 * Promote `<w:fldSimple w:instr=" PAGE \\* MERGEFORMAT "><w:r>…</w:r></w:fldSimple>`
 * (and the equivalent `NUMPAGES` form) into a typed {@link Run}
 * carrying a single {@link PageNumberFieldLeaf}.
 *
 * Only the simple PAGE / NUMPAGES forms are promoted — anything
 * else (TOC, REF, MERGEFIELD, …) keeps the existing opaque-inline
 * round-trip path, which is byte-stable.
 *
 * The serializer (`serializeRunOrFieldWrapper`) already reverses
 * this promotion: a `Run` whose only child is a
 * `PageNumberFieldLeaf` re-emits as a `<w:fldSimple>` wrapper, so
 * the parse → serialize → parse cycle stays loss-free.
 */
function tryParseFldSimpleAsPageNumber(
  entry: Record<string, unknown>,
  mintNodeId: IdMinter
): Run | null {
  const instrRaw = attrOf(entry, "w:instr");
  if (!instrRaw) return null;
  const field = detectPageNumberField(instrRaw);
  if (!field) return null;
  const children = (entry["w:fldSimple"] as unknown[] | undefined) ?? [];
  const innerRun = findElementEntry(children, "w:r");
  let runProps: RunProperties = {};
  let cachedText: string | undefined;
  if (innerRun) {
    const runChildren = (innerRun["w:r"] as unknown[] | undefined) ?? [];
    const rPr = findElementEntry(runChildren, "w:rPr");
    if (rPr) runProps = parseRunProperties(rPr);
    for (const c of elementEntries(runChildren)) {
      if (ooxml.getTag(c) === "w:t") {
        const v = (c["w:t"] as unknown[] | undefined) ?? [];
        const txt = v
          .map((n) => (typeof n === "object" && n && "#text" in n ? String((n as { "#text": unknown })["#text"] ?? "") : ""))
          .join("");
        if (txt.length > 0) cachedText = txt;
      }
    }
  }
  const leaf: PageNumberFieldLeaf = {
    kind: "page-number-field",
    id: mintNodeId(),
    field,
    instr: instrRaw,
    ...(cachedText !== undefined ? { cachedText } : {}),
  };
  return {
    kind: "run",
    id: mintNodeId(),
    properties: runProps,
    children: [leaf],
  };
}

/**
 * Promote `<w:r><w:commentReference w:id="N"/></w:r>` into the
 * typed {@link CommentReference} inline node. The serializer
 * (`serializeInline`) inverses this by re-emitting the
 * `<w:r>` wrapper. We only promote when the run contains exactly
 * the comment-reference element (no `<w:rPr>` properties, no
 * sibling text/breaks/etc.) so that authored runs that legitimately
 * mix a reference with other run content stay opaque.
 */
function tryParseRunAsCommentReference(
  entry: Record<string, unknown>,
  mintNodeId: IdMinter
): CommentReference | null {
  const children = (entry["w:r"] as unknown[] | undefined) ?? [];
  const elementChildren = elementEntries(children);
  if (elementChildren.length !== 1) return null;
  const only = elementChildren[0]!;
  if (ooxml.getTag(only) !== "w:commentReference") return null;
  const commentId = attrOf(only, "w:id");
  if (commentId === undefined) return null;
  return {
    kind: "comment-reference",
    id: mintNodeId(),
    commentId,
  };
}

/**
 * Recognise the field instruction string used by `<w:fldSimple>`.
 * Word writes the instruction with surrounding spaces and an
 * optional `\* MERGEFORMAT` (or other) switch — we tokenise and
 * return the canonical name when it matches a supported field.
 */
function detectPageNumberField(instr: string): "PAGE" | "NUMPAGES" | null {
  const tokens = instr.trim().split(/\s+/);
  const head = tokens[0]?.toUpperCase();
  if (head === "PAGE") return "PAGE";
  if (head === "NUMPAGES") return "NUMPAGES";
  return null;
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

/** See `parseParagraphProperties` for why this is exported. */
export function parseRunProperties(entry: Record<string, unknown>): RunProperties {
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
        const hAnsi = attrOf(c, "w:hAnsi");
        if (hAnsi) props.fontFamilyHAnsi = hAnsi;
        const eastAsia = attrOf(c, "w:eastAsia");
        if (eastAsia) props.fontFamilyEastAsia = eastAsia;
        const cs = attrOf(c, "w:cs");
        if (cs) props.fontFamilyComplexScript = cs;
        // Theme refs are projected into typed fields so the cascade
        // resolver can consult them. The raw element is still pushed
        // into `opaqueProps` for byte-identical round-trip — the
        // serializer suppresses it when an explicit `fontFamily` has
        // been set on the run, otherwise emits it verbatim. See
        // `agent/style-resolver.ts` for the projection rule (literal
        // ascii wins over asciiTheme when both are present).
        const asciiTheme = attrOf(c, "w:asciiTheme");
        if (asciiTheme) props.fontFamilyAsciiTheme = asciiTheme;
        const hAnsiTheme = attrOf(c, "w:hAnsiTheme");
        if (hAnsiTheme) props.fontFamilyHAnsiTheme = hAnsiTheme;
        const eastAsiaTheme = attrOf(c, "w:eastAsiaTheme");
        if (eastAsiaTheme) props.fontFamilyEastAsiaTheme = eastAsiaTheme;
        const csTheme = attrOf(c, "w:cstheme");
        if (csTheme) props.fontFamilyComplexScriptTheme = csTheme;
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
      if (t === "page") {
        return { kind: "page-break", id: mintNodeId() };
      }
      const breakType = t === "column" || t === "textWrapping" ? t : undefined;
      return { kind: "break", id: mintNodeId(), ...(breakType ? { breakType } : {}) };
    }
    case "w:lastRenderedPageBreak":
      return { kind: "last-rendered-page-break", id: mintNodeId() };
    case "w:tab":
      return { kind: "tab", id: mintNodeId() };
    case "w:drawing":
      return parseDrawing(entry, mintNodeId, currentChartResolver);
    case "w:object": {
      const ole = parseEmbeddedSpreadsheet(entry, mintNodeId, currentDocRelResolver);
      if (ole) return ole;
      return { kind: "opaque", id: mintNodeId(), raw: captureOpaque(entry) };
    }
    default:
      return { kind: "opaque", id: mintNodeId(), raw: captureOpaque(entry) };
  }
}

/**
 * Try to recognise a `<w:object>` run-child as an OLE-embedded
 * Excel.Sheet workbook. Drills into both the simple form
 * (`<w:object><o:OLEObject .../></w:object>`) and the
 * `<mc:AlternateContent>` wrapper modern Word emits. Returns `null`
 * for non-Excel OLE objects (Word, PowerPoint, generic CFB blobs)
 * so the caller falls back to opaque preservation.
 */
function parseEmbeddedSpreadsheet(
  entry: Record<string, unknown>,
  mintNodeId: IdMinter,
  resolveRel: ((relId: string) => string | undefined) | undefined
): EmbeddedSpreadsheet | null {
  if (!resolveRel) return null;
  const objectChildren = (entry["w:object"] as unknown[] | undefined) ?? [];
  let oleEntry = findElementEntry(objectChildren, "o:OLEObject");
  if (!oleEntry) {
    const ac = findElementEntry(objectChildren, "mc:AlternateContent");
    if (ac) {
      const choice =
        findElementEntry((ac["mc:AlternateContent"] as unknown[] | undefined) ?? [], "mc:Choice") ??
        findElementEntry((ac["mc:AlternateContent"] as unknown[] | undefined) ?? [], "mc:Fallback");
      if (choice) {
        const innerChildren = (choice[ooxml.getTag(choice)] as unknown[] | undefined) ?? [];
        oleEntry = findElementEntry(innerChildren, "o:OLEObject");
      }
    }
  }
  if (!oleEntry) return null;
  const oleAttrs: Record<string, string> = {};
  const oleAttrBag = oleEntry[":@"];
  if (oleAttrBag && typeof oleAttrBag === "object") {
    for (const [k, v] of Object.entries(oleAttrBag as Record<string, unknown>)) {
      const name = k.startsWith("@_") ? k.slice(2) : k;
      oleAttrs[name] = String(v);
    }
  }
  const progId = oleAttrs["ProgID"] ?? "";
  if (!progId.startsWith("Excel.Sheet")) return null;
  const oleRelId = oleAttrs["r:id"] ?? "";
  if (!oleRelId) return null;
  const embeddingPartPath = resolveRel(oleRelId);
  if (!embeddingPartPath) return null;
  const embeddingKind: "xlsx" | "bin" = embeddingPartPath.toLowerCase().endsWith(".xlsx")
    ? "xlsx"
    : "bin";

  // Optional VML preview: <v:shape><v:imagedata r:id="rIdImage" .../>.
  let previewImageRelId: string | undefined;
  let previewImagePartPath: string | undefined;
  for (const c of elementEntries(objectChildren)) {
    if (ooxml.getTag(c) !== "v:shape") continue;
    const shapeChildren = (c["v:shape"] as unknown[] | undefined) ?? [];
    const imagedata = findElementEntry(shapeChildren, "v:imagedata");
    if (!imagedata) continue;
    const rid = attrOf(imagedata, "r:id");
    if (rid) {
      previewImageRelId = rid;
      previewImagePartPath = resolveRel(rid);
    }
    break;
  }

  return {
    kind: "embedded-spreadsheet",
    id: mintNodeId(),
    oleRelId,
    embeddingPartPath,
    progId,
    embeddingKind,
    ...(previewImageRelId ? { previewImageRelId } : {}),
    ...(previewImagePartPath ? { previewImagePartPath } : {}),
    oleObjectAttrs: oleAttrs,
    raw: captureOpaque(entry),
  };
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

/**
 * Decode the OOXML "toggle" boolean encoding for `<w:keepNext/>`,
 * `<w:keepLines/>`, `<w:bold/>`, `<w:widowControl/>`, etc. The element's
 * presence alone means `true`; an explicit `w:val="0"` (or `"false"` /
 * `"off"`) flips it to `false`. Used by the paragraph-property parser
 * to surface pagination flags as typed booleans.
 */
function ooxmlBoolFlag(entry: Record<string, unknown>): boolean {
  const v = attrOf(entry, "w:val");
  return boolAttr(v, true);
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
