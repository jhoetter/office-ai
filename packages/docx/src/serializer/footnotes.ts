import { ooxml } from "@officeai/core";
import type { BlockNode, DocxSnapshot, Footnote, FootnotesPart } from "../model/types.js";
import { opaqueToEntry } from "../parser/xml-helpers.js";

/**
 * Serializer for `word/footnotes.xml` (F1).
 *
 * Byte-preservation contract:
 *   - When `dirty.footnotes` is `false`, the part is left alone — the
 *     cloned container already carries the original bytes, so the
 *     round-trip is byte-identical (and we never even visit this code
 *     path on no-touch saves).
 *   - When the flag is set, every footnote with a `raw` envelope
 *     re-emits from that envelope verbatim; only footnotes whose
 *     mutating command dropped `raw` are regenerated from the typed
 *     model. That is what guarantees a single-footnote edit only
 *     touches that one footnote's bytes.
 *
 * We also (re)register the relationship in
 * `word/_rels/document.xml.rels` and the `[Content_Types].xml`
 * override the first time the part materialises in the package, so
 * Word picks up footnotes added to a fresh document on save.
 */

const FOOTNOTES_PART = "word/footnotes.xml";
const FOOTNOTES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
const FOOTNOTES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const MAIN_PART = "word/document.xml";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const DEFAULT_FOOTNOTES_ROOT_ATTRS: Record<string, string> = {
  "xmlns:w": W_NS,
};

export function serializeFootnotesPart(
  container: ooxml.OoxmlContainer,
  snapshot: DocxSnapshot,
  serializeBlock: (block: BlockNode) => unknown
): void {
  if (!snapshot.dirty.footnotes) return;
  const part = snapshot.root.footnotesPart;
  if (!part || part.footnotes.length === 0) {
    if (container.has(FOOTNOTES_PART)) {
      removeFootnotesPart(container);
    }
    return;
  }

  ensureFootnotesPart(container);
  const xml = serializeFootnotesXml(part, serializeBlock);
  container.writeText(FOOTNOTES_PART, xml);
}

function serializeFootnotesXml(part: FootnotesPart, serializeBlock: (block: BlockNode) => unknown): string {
  const children: unknown[] = [];
  for (const fn of part.footnotes) {
    children.push(serializeFootnote(fn, serializeBlock));
  }
  if (part.tail) {
    for (const tail of part.tail) {
      children.push(opaqueToEntry(tail));
    }
  }
  const rootAttrs: Record<string, string> = {};
  const merged = { ...DEFAULT_FOOTNOTES_ROOT_ATTRS, ...part.rootAttrs };
  for (const [k, v] of Object.entries(merged)) {
    rootAttrs[`@_${k}`] = v;
  }
  const tree = [{ "w:footnotes": children, ":@": rootAttrs }];
  return ooxml.serializeXml(tree, { xmlDeclaration: XML_DECL });
}

/**
 * Per-footnote emission: when the footnote still carries its `raw`
 * envelope (i.e. it has not been touched since parse), re-emit the
 * cached subtree verbatim — this is what keeps every untouched
 * footnote's bytes byte-identical even when a sibling was edited.
 *
 * When `raw` is absent (the footnote was created or replaced by a
 * mutating command), regenerate the `<w:footnote>` envelope from the
 * typed model.
 */
function serializeFootnote(fn: Footnote, serializeBlock: (block: BlockNode) => unknown): unknown {
  if (fn.raw) {
    return opaqueToEntry(fn.raw);
  }
  const attrs: Record<string, string> = {
    "@_w:id": String(fn.id),
  };
  if (fn.type !== "normal") {
    attrs["@_w:type"] = fn.type;
  }
  const body = fn.body.map((b) => serializeBlock(b));
  return { "w:footnote": body, ":@": attrs };
}

function ensureFootnotesPart(container: ooxml.OoxmlContainer): void {
  if (!container.has(FOOTNOTES_PART)) {
    const empty = ooxml.serializeXml([{ "w:footnotes": [], ":@": { "@_xmlns:w": W_NS } }], {
      xmlDeclaration: XML_DECL,
    });
    container.addPart(FOOTNOTES_PART, new TextEncoder().encode(empty));
  }

  const rels = ooxml.RelationshipGraph.loadFor(container, MAIN_PART);
  if (rels.byType(FOOTNOTES_REL_TYPE).length === 0) {
    rels.add({ type: FOOTNOTES_REL_TYPE, target: "footnotes.xml" });
    rels.writeBack(container);
  }

  const ct = ooxml.ContentTypes.load(container);
  if (!ct.hasOverride("/word/footnotes.xml")) {
    ct.addOverride("/word/footnotes.xml", FOOTNOTES_CONTENT_TYPE);
    ct.writeBack(container);
  }
}

function removeFootnotesPart(container: ooxml.OoxmlContainer): void {
  if (!container.has(FOOTNOTES_PART)) return;
  container.removePart(FOOTNOTES_PART);
  const rels = ooxml.RelationshipGraph.loadFor(container, MAIN_PART);
  for (const r of rels.byType(FOOTNOTES_REL_TYPE)) {
    rels.remove(r.id);
  }
  rels.writeBack(container);
  const ct = ooxml.ContentTypes.load(container);
  ct.removeOverride("/word/footnotes.xml");
  ct.writeBack(container);
}
