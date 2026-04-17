import { ooxml } from "@officeai/core";
import type {
  BlockNode,
  DocxComment,
  DocxDocument,
  DocxSnapshot,
  Hyperlink,
  InlineNode,
  Paragraph,
  ParagraphProperties,
  RevisionWrapper,
  Run,
  RunChild,
  RunProperties,
} from "../model/types.js";
import { opaqueToEntry } from "../parser/xml-helpers.js";
import { DocxSerializeError } from "./errors.js";

const MAIN_PART = "word/document.xml";
const COMMENTS_PART = "word/comments.xml";
const COMMENTS_EXTENDED_PART = "word/commentsExtended.xml";
const COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_EXTENDED_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const COMMENTS_EXTENDED_CONTENT_TYPE = "application/vnd.ms-word.commentsExtended+xml";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const COMMENTS_ROOT_ATTRS: Record<string, string> = {
  "@_xmlns:w": W_NS,
  "@_xmlns:w14": W14_NS,
};

const COMMENTS_EXTENDED_ROOT_ATTRS: Record<string, string> = {
  "@_xmlns:w15": W15_NS,
};

export async function serializeDocx(snapshot: DocxSnapshot): Promise<ArrayBuffer> {
  const container = snapshot.container.clone();

  if (snapshot.dirty.body) {
    try {
      const xml = serializeDocumentXml(snapshot.root);
      container.writeText(MAIN_PART, xml);
    } catch (err) {
      throw new DocxSerializeError("body-failed", "Failed to serialize document.xml", { cause: err });
    }
  }

  if (snapshot.dirty.comments) {
    if (snapshot.root.comments.length > 0) {
      ensureCommentsPart(container);
      const xml = serializeCommentsXml(snapshot.root.comments);
      container.writeText(COMMENTS_PART, xml);
    } else if (container.has(COMMENTS_PART)) {
      removeCommentsPart(container);
    }
  }

  // commentsExtended.xml. We re-emit it whenever EITHER the comments part or
  // the extended part is dirty, because the two are tightly coupled (e.g.
  // resolving a comment dirties only `commentsExtended`, but adding a comment
  // dirties `comments` and may also need a fresh extended record).
  if (snapshot.dirty.commentsExtended || snapshot.dirty.comments) {
    const records = collectCommentsExtended(snapshot.root.comments);
    if (records.length > 0) {
      ensureCommentsExtendedPart(container);
      const xml = serializeCommentsExtendedXml(records);
      container.writeText(COMMENTS_EXTENDED_PART, xml);
    } else if (container.has(COMMENTS_EXTENDED_PART)) {
      removeCommentsExtendedPart(container);
    }
  }

  return container.serialize();
}

function serializeDocumentXml(doc: DocxDocument): string {
  const bodyChildren: unknown[] = [];
  for (const block of doc.body) {
    bodyChildren.push(serializeBlock(block));
  }
  const bodyEntry: Record<string, unknown> = { "w:body": bodyChildren };
  const docEntry: Record<string, unknown> = {
    "w:document": [bodyEntry],
  };
  const rootAttrs: Record<string, string> = {};
  for (const [k, v] of Object.entries(doc.documentRootAttrs)) {
    rootAttrs[`@_${k}`] = v;
  }
  if (Object.keys(rootAttrs).length > 0) {
    docEntry[":@"] = rootAttrs;
  }
  const tree = [docEntry];
  const xml = ooxml.serializeXml(tree, { xmlDeclaration: XML_DECL });
  return xml;
}

function serializeBlock(block: BlockNode): unknown {
  switch (block.kind) {
    case "paragraph":
      return serializeParagraph(block);
    case "table":
    case "section-break":
    case "opaque-block":
      return opaqueToEntry(block.raw);
    default: {
      const _exhaustive: never = block;
      throw new DocxSerializeError("unknown-block", `Unknown block kind: ${JSON.stringify(_exhaustive)}`);
    }
  }
}

function serializeParagraph(p: Paragraph): unknown {
  const children: unknown[] = [];
  const pPr = serializeParagraphProperties(p.properties);
  if (pPr) children.push(pPr);
  for (const c of p.children) {
    children.push(serializeInline(c));
  }
  return { "w:p": children };
}

function serializeParagraphProperties(props: ParagraphProperties): unknown | null {
  const out: unknown[] = [];
  if (props.styleId) {
    out.push(makeEl("w:pStyle", { "w:val": props.styleId }));
  }
  if (props.numbering) {
    // Numbering is also represented in opaqueProps to preserve full attrs;
    // we do NOT emit a typed numPr here to avoid duplication.
  }
  if (props.indentation) {
    const ind: Record<string, string> = {};
    if (props.indentation.left !== undefined) ind["w:left"] = String(props.indentation.left);
    if (props.indentation.right !== undefined) ind["w:right"] = String(props.indentation.right);
    if (props.indentation.firstLine !== undefined) ind["w:firstLine"] = String(props.indentation.firstLine);
    if (props.indentation.hanging !== undefined) ind["w:hanging"] = String(props.indentation.hanging);
    if (Object.keys(ind).length > 0) out.push(makeEl("w:ind", ind));
  }
  if (props.spacing) {
    const sp: Record<string, string> = {};
    if (props.spacing.before !== undefined) sp["w:before"] = String(props.spacing.before);
    if (props.spacing.after !== undefined) sp["w:after"] = String(props.spacing.after);
    if (props.spacing.line !== undefined) sp["w:line"] = String(props.spacing.line);
    if (props.spacing.lineRule) sp["w:lineRule"] = props.spacing.lineRule;
    if (Object.keys(sp).length > 0) out.push(makeEl("w:spacing", sp));
  }
  if (props.alignment) {
    const map: Record<string, string> = {
      left: "left",
      center: "center",
      right: "right",
      justify: "both",
    };
    out.push(makeEl("w:jc", { "w:val": map[props.alignment] ?? props.alignment }));
  }
  if (props.opaqueProps) {
    for (const o of props.opaqueProps) {
      out.push(opaqueToEntry(o));
    }
  }
  if (out.length === 0) return null;
  return { "w:pPr": out };
}

function serializeInline(node: InlineNode): unknown {
  switch (node.kind) {
    case "run":
      return serializeRun(node);
    case "hyperlink":
      return serializeHyperlink(node);
    case "comment-range-start":
      return makeEl("w:commentRangeStart", { "w:id": node.commentId });
    case "comment-range-end":
      return makeEl("w:commentRangeEnd", { "w:id": node.commentId });
    case "comment-reference":
      return { "w:r": [makeEl("w:commentReference", { "w:id": node.commentId })] };
    case "revision":
      return serializeRevisionWrapper(node);
    case "opaque-inline":
      return opaqueToEntry(node.raw);
    default: {
      const _exhaustive: never = node;
      throw new DocxSerializeError("unknown-inline", `Unknown inline kind: ${JSON.stringify(_exhaustive)}`);
    }
  }
}

function serializeRun(r: Run): unknown {
  const children: unknown[] = [];
  const rPr = serializeRunProperties(r.properties);
  if (rPr) children.push(rPr);
  for (const c of r.children) {
    children.push(serializeRunChild(c));
  }
  return { "w:r": children };
}

function serializeRunProperties(props: RunProperties): unknown | null {
  const out: unknown[] = [];
  if (props.fontFamily) {
    out.push(makeEl("w:rFonts", { "w:ascii": props.fontFamily, "w:hAnsi": props.fontFamily }));
  }
  if (props.bold !== undefined) {
    out.push(props.bold ? makeEl("w:b") : makeEl("w:b", { "w:val": "false" }));
  }
  if (props.italic !== undefined) {
    out.push(props.italic ? makeEl("w:i") : makeEl("w:i", { "w:val": "false" }));
  }
  if (props.strike !== undefined && props.strike) {
    out.push(makeEl("w:strike"));
  }
  if (props.underline !== undefined) {
    if (props.underline === false) {
      out.push(makeEl("w:u", { "w:val": "none" }));
    } else if (props.underline === true) {
      out.push(makeEl("w:u", { "w:val": "single" }));
    } else {
      out.push(makeEl("w:u", { "w:val": props.underline }));
    }
  }
  if (props.color) {
    out.push(makeEl("w:color", { "w:val": props.color }));
  }
  if (props.fontSize !== undefined) {
    out.push(makeEl("w:sz", { "w:val": String(props.fontSize) }));
    out.push(makeEl("w:szCs", { "w:val": String(props.fontSize) }));
  }
  if (props.highlight) {
    out.push(makeEl("w:highlight", { "w:val": props.highlight }));
  }
  if (props.opaqueProps) {
    for (const o of props.opaqueProps) {
      // Skip opaque rFonts since we already emitted a typed one if fontFamily set
      if (o.tag === "w:rFonts" && props.fontFamily) continue;
      out.push(opaqueToEntry(o));
    }
  }
  if (out.length === 0) return null;
  return { "w:rPr": out };
}

function serializeRunChild(c: RunChild): unknown {
  switch (c.kind) {
    case "text": {
      const tag = c.isDelText ? "w:delText" : "w:t";
      const attrs: Record<string, string> = {};
      const needsPreserve = c.xmlSpacePreserve || /^\s|\s$/.test(c.text) || c.text === "";
      if (needsPreserve) attrs["xml:space"] = "preserve";
      const entry: Record<string, unknown> = { [tag]: [{ "#text": c.text }] };
      if (Object.keys(attrs).length > 0) {
        entry[":@"] = Object.fromEntries(Object.entries(attrs).map(([k, v]) => [`@_${k}`, v]));
      }
      return entry;
    }
    case "break":
      return c.breakType ? makeEl("w:br", { "w:type": c.breakType }) : { "w:br": [] };
    case "tab":
      return { "w:tab": [] };
    case "drawing":
    case "opaque":
      return opaqueToEntry(c.raw);
    default: {
      const _exhaustive: never = c;
      throw new DocxSerializeError(
        "unknown-run-child",
        `Unknown run child kind: ${JSON.stringify(_exhaustive)}`
      );
    }
  }
}

function serializeHyperlink(h: Hyperlink): unknown {
  const children = h.children.map((r) => serializeRun(r));
  const attrs: Record<string, string> = {};
  if (h.relationshipId) attrs["r:id"] = h.relationshipId;
  if (h.anchor) attrs["w:anchor"] = h.anchor;
  return makeEntry("w:hyperlink", children, attrs);
}

function serializeRevisionWrapper(rev: RevisionWrapper): unknown {
  const tag = rev.revisionType === "ins" ? "w:ins" : "w:del";
  const children = rev.children.map(serializeInline);
  const attrs: Record<string, string> = {
    "w:id": rev.revisionId,
    "w:author": rev.author,
    "w:date": rev.date,
  };
  return makeEntry(tag, children, attrs);
}

function serializeCommentsXml(comments: ReadonlyArray<DocxComment>): string {
  const children: unknown[] = comments.map((c) => serializeComment(c));
  const tree = [{ "w:comments": children, ":@": COMMENTS_ROOT_ATTRS }];
  return ooxml.serializeXml(tree, { xmlDeclaration: XML_DECL });
}

function serializeComment(c: DocxComment): unknown {
  const paraId = effectiveParaId(c);
  const body: unknown[] = c.body.map((b, i) => {
    const entry = serializeBlock(b);
    if (i === 0 && b.kind === "paragraph") return injectParaId(entry, paraId);
    return entry;
  });
  const attrs: Record<string, string> = {
    "w:id": c.id,
    "w:author": c.author,
    "w:date": c.date,
  };
  if (c.initials) attrs["w:initials"] = c.initials;
  return makeEntry("w:comment", body, attrs);
}

/**
 * Inject (or override) `w14:paraId` on a `w:p` entry so that
 * `commentsExtended.xml` can reference it. The attribute lives on the `:@`
 * sibling object; we mutate a shallow copy of the entry to keep the input
 * immutable.
 */
function injectParaId(entry: unknown, paraId: string): unknown {
  if (!entry || typeof entry !== "object") return entry;
  const e = entry as Record<string, unknown>;
  const tag = Object.keys(e).find((k) => k !== ":@");
  if (tag !== "w:p") return entry;
  const next: Record<string, unknown> = { ...e };
  const existingAttrs = (e[":@"] as Record<string, string> | undefined) ?? {};
  next[":@"] = { ...existingAttrs, "@_w14:paraId": paraId };
  return next;
}

interface CommentExtendedRecord {
  paraId: string;
  done: boolean;
  parentParaId?: string;
}

/**
 * Materialize the records that `commentsExtended.xml` needs. Only comments
 * that are resolved or that belong to a thread (parent or child) get a
 * record — top-level, open comments don't need one and Word omits them too.
 */
function collectCommentsExtended(comments: ReadonlyArray<DocxComment>): CommentExtendedRecord[] {
  const idToParaId = new Map<string, string>();
  for (const c of comments) idToParaId.set(c.id, effectiveParaId(c));

  const childIds = new Set<string>();
  for (const c of comments) {
    if (c.parentId) childIds.add(c.parentId);
  }

  const out: CommentExtendedRecord[] = [];
  for (const c of comments) {
    const isResolved = c.resolved === true;
    const isReply = c.parentId !== undefined;
    const isThreadParent = childIds.has(c.id);
    if (!isResolved && !isReply && !isThreadParent) continue;
    const rec: CommentExtendedRecord = {
      paraId: effectiveParaId(c),
      done: isResolved,
    };
    if (c.parentId) {
      const parent = idToParaId.get(c.parentId);
      if (parent) rec.parentParaId = parent;
    }
    out.push(rec);
  }
  return out;
}

function serializeCommentsExtendedXml(records: ReadonlyArray<CommentExtendedRecord>): string {
  const children: unknown[] = records.map((r) => {
    const attrs: Record<string, string> = {
      "w15:paraId": r.paraId,
      "w15:done": r.done ? "1" : "0",
    };
    if (r.parentParaId) attrs["w15:parentPaIdRef"] = r.parentParaId;
    return makeEntry("w15:commentEx", [], attrs);
  });
  const tree = [{ "w15:commentsEx": children, ":@": COMMENTS_EXTENDED_ROOT_ATTRS }];
  return ooxml.serializeXml(tree, { xmlDeclaration: XML_DECL });
}

/**
 * Return the comment's persisted `paraId` if it has one, otherwise mint a
 * deterministic 8-hex-char id from the comment id. Determinism matters so
 * that round-tripping the same in-memory snapshot twice produces identical
 * bytes — which is the whole point of byte-preservation in this codebase.
 */
function effectiveParaId(c: DocxComment): string {
  if (c.paraId) return c.paraId;
  return deriveParaIdFromCommentId(c.id);
}

function deriveParaIdFromCommentId(commentId: string): string {
  const h = hashString(`officeai/comment-paraid/${commentId}`);
  return h.slice(0, 8).toUpperCase();
}

function hashString(s: string): string {
  // Tiny FNV-1a 32-bit, then expand to 8 hex chars. Good enough as a
  // deterministic non-colliding-in-practice id for the sizes we're dealing
  // with (Word documents have at most thousands of comments, FNV-1a 32-bit
  // is fine; we're not using this for security).
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  const u = h >>> 0;
  return u.toString(16).padStart(8, "0");
}

function ensureCommentsPart(container: ooxml.OoxmlContainer): void {
  if (!container.has(COMMENTS_PART)) {
    const empty = ooxml.serializeXml([{ "w:comments": [], ":@": COMMENTS_ROOT_ATTRS }], {
      xmlDeclaration: XML_DECL,
    });
    container.addPart(COMMENTS_PART, new TextEncoder().encode(empty));
  }

  // Ensure the relationship exists in word/_rels/document.xml.rels.
  const rels = ooxml.RelationshipGraph.loadFor(container, MAIN_PART);
  const exists = rels.byType(COMMENTS_REL_TYPE).length > 0;
  if (!exists) {
    rels.add({ type: COMMENTS_REL_TYPE, target: "comments.xml" });
    rels.writeBack(container);
  }

  // Ensure [Content_Types].xml has the override.
  const ct = ooxml.ContentTypes.load(container);
  if (!ct.hasOverride("/word/comments.xml")) {
    ct.addOverride("/word/comments.xml", COMMENTS_CONTENT_TYPE);
    ct.writeBack(container);
  }
}

function removeCommentsPart(container: ooxml.OoxmlContainer): void {
  container.removePart(COMMENTS_PART);
  const rels = ooxml.RelationshipGraph.loadFor(container, MAIN_PART);
  for (const r of rels.byType(COMMENTS_REL_TYPE)) {
    rels.remove(r.id);
  }
  rels.writeBack(container);
  const ct = ooxml.ContentTypes.load(container);
  ct.removeOverride("/word/comments.xml");
  ct.writeBack(container);
}

function ensureCommentsExtendedPart(container: ooxml.OoxmlContainer): void {
  if (!container.has(COMMENTS_EXTENDED_PART)) {
    const empty = ooxml.serializeXml([{ "w15:commentsEx": [], ":@": COMMENTS_EXTENDED_ROOT_ATTRS }], {
      xmlDeclaration: XML_DECL,
    });
    container.addPart(COMMENTS_EXTENDED_PART, new TextEncoder().encode(empty));
  }
  const rels = ooxml.RelationshipGraph.loadFor(container, MAIN_PART);
  if (rels.byType(COMMENTS_EXTENDED_REL_TYPE).length === 0) {
    rels.add({ type: COMMENTS_EXTENDED_REL_TYPE, target: "commentsExtended.xml" });
    rels.writeBack(container);
  }
  const ct = ooxml.ContentTypes.load(container);
  if (!ct.hasOverride("/word/commentsExtended.xml")) {
    ct.addOverride("/word/commentsExtended.xml", COMMENTS_EXTENDED_CONTENT_TYPE);
    ct.writeBack(container);
  }
}

function removeCommentsExtendedPart(container: ooxml.OoxmlContainer): void {
  if (!container.has(COMMENTS_EXTENDED_PART)) return;
  container.removePart(COMMENTS_EXTENDED_PART);
  const rels = ooxml.RelationshipGraph.loadFor(container, MAIN_PART);
  for (const r of rels.byType(COMMENTS_EXTENDED_REL_TYPE)) {
    rels.remove(r.id);
  }
  rels.writeBack(container);
  const ct = ooxml.ContentTypes.load(container);
  ct.removeOverride("/word/commentsExtended.xml");
  ct.writeBack(container);
}

/* ── small builders ──────────────────────────────────────────────────────── */

function makeEl(tag: string, attrs?: Record<string, string>): unknown {
  return makeEntry(tag, [], attrs ?? {});
}

function makeEntry(tag: string, children: unknown[], attrs: Record<string, string>): unknown {
  const entry: Record<string, unknown> = { [tag]: children };
  if (Object.keys(attrs).length > 0) {
    const out: Record<string, string> = {};
    for (const [k, v] of Object.entries(attrs)) {
      out[`@_${k}`] = v;
    }
    entry[":@"] = out;
  }
  return entry;
}
