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
const COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const COMMENTS_ROOT_ATTRS: Record<string, string> = {
  "@_xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
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
  const body: unknown[] = c.body.map((b) => serializeBlock(b));
  const attrs: Record<string, string> = {
    "w:id": c.id,
    "w:author": c.author,
    "w:date": c.date,
  };
  if (c.initials) attrs["w:initials"] = c.initials;
  return makeEntry("w:comment", body, attrs);
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
