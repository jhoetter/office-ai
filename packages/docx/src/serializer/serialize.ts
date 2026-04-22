import { ooxml } from "@officeai/core";
import type {
  BlockNode,
  DocxComment,
  DocxDocument,
  DocxSnapshot,
  Hyperlink,
  InlineNode,
  OpaqueXml,
  Paragraph,
  ParagraphProperties,
  RevisionWrapper,
  Run,
  RunChild,
  RunProperties,
} from "../model/types.js";
import { ATTR_KEY, opaqueToEntry } from "../parser/xml-helpers.js";
import { serializeChartDrawing, serializeChartParts } from "./charts.js";
import { serializeEmbeddingParts } from "./embeddings.js";
import { DocxSerializeError } from "./errors.js";
import { serializeFootnotesPart } from "./footnotes.js";
import { serializeHeaderFooterParts } from "./headers-footers.js";
import { serializeInlineImageDrawing } from "./images.js";
import { serializeMediaParts } from "./media.js";
import { serializeNumberingPart } from "./numbering.js";
import { serializeRelationshipsParts } from "./relationships.js";
import { serializeSectionProperties } from "./sections.js";
import { serializeTable, serializeTableFromRaw } from "./tables.js";

const MAIN_PART = "word/document.xml";
const COMMENTS_PART = "word/comments.xml";
const COMMENTS_EXTENDED_PART = "word/commentsExtended.xml";
const COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_EXTENDED_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const COMMENTS_EXTENDED_CONTENT_TYPE = "application/vnd.ms-word.commentsExtended+xml";
const NUMBERING_PART = "word/numbering.xml";
const NUMBERING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
const NUMBERING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

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

  // Headers / footers. The headers-footers serializer skips parts whose
  // path is not in the dirty set, which is what guarantees byte-identity for
  // untouched parts on a no-touch round-trip.
  try {
    serializeHeaderFooterParts(container, snapshot, serializeBlock);
  } catch (err) {
    if (err instanceof DocxSerializeError) throw err;
    throw new DocxSerializeError("header-footer-failed", "Failed to serialize a header/footer part", {
      cause: err,
    });
  }

  // Media bytes. Same dirty-set-only contract as headers/footers — only
  // parts the snapshot reports as dirty are re-written.
  try {
    serializeMediaParts(container, snapshot);
  } catch (err) {
    if (err instanceof DocxSerializeError) throw err;
    throw new DocxSerializeError("media-failed", "Failed to write media parts", { cause: err });
  }

  // Relationships parts. Rewritten from the typed model when dirty.
  try {
    serializeRelationshipsParts(container, snapshot);
  } catch (err) {
    if (err instanceof DocxSerializeError) throw err;
    throw new DocxSerializeError("rels-failed", "Failed to write relationships parts", { cause: err });
  }

  // Chart parts: regenerate `word/charts/chartN.xml` from the typed
  // model and (re)materialise the embedded xlsx package + its
  // per-chart relationships so Office's "Edit Data" UI sees a real,
  // round-trippable workbook rather than a dangling reference.
  try {
    await serializeChartParts(container, snapshot);
  } catch (err) {
    if (err instanceof DocxSerializeError) throw err;
    throw new DocxSerializeError("chart-failed", "Failed to write chart parts", { cause: err });
  }

  // Embedded binary parts (`word/embeddings/*.xlsx`, …). Authored by
  // OLE-Excel-spreadsheet inserts and re-materialised here so Office
  // sees the live `.xlsx` package on double-click. Untouched embeds
  // ride the container's part cache; only entries in `dirty.embeddings`
  // are re-written.
  try {
    await serializeEmbeddingParts(container, snapshot);
  } catch (err) {
    if (err instanceof DocxSerializeError) throw err;
    throw new DocxSerializeError("embedding-failed", "Failed to write embedding parts", { cause: err });
  }

  // Footnotes part (`word/footnotes.xml`). Untouched documents leave
  // the byte cache alone (this code path only activates when
  // `dirty.footnotes` is true). When dirty, every footnote with a
  // cached `raw` envelope re-emits byte-identical, so a single-footnote
  // edit only re-writes that footnote's bytes.
  try {
    serializeFootnotesPart(container, snapshot, serializeBlock);
  } catch (err) {
    if (err instanceof DocxSerializeError) throw err;
    throw new DocxSerializeError("footnotes-failed", "Failed to write footnotes.xml", { cause: err });
  }

  // Numbering definitions (`word/numbering.xml`). Skipped unless
  // `dirty.numbering` is set, so the part round-trips byte-identical
  // when no command has touched its definitions. B7 wires this to
  // `docx:set-paragraph-list` so the first list applied to a doc
  // without `word/numbering.xml` materialises the part, registers a
  // relationship + Content_Types override, and emits the typed
  // `<w:abstractNum>` / `<w:num>` carrier from the snapshot.
  try {
    if (snapshot.dirty.numbering && snapshot.root.numbering) {
      ensureNumberingPart(container);
    }
    serializeNumberingPart(container, snapshot);
  } catch (err) {
    if (err instanceof DocxSerializeError) throw err;
    throw new DocxSerializeError("numbering-failed", "Failed to write numbering.xml", { cause: err });
  }

  // [Content_Types].xml. We update this when the caller has set the
  // contentTypes dirty flag (currently driven by `docx:insert-image` when
  // it adds a new image MIME type or registers a new override). The typed
  // model for content types is the union of every media MIME we know
  // about plus the existing defaults/overrides.
  if (snapshot.dirty.contentTypes) {
    try {
      ensureMediaContentTypes(container, snapshot);
    } catch (err) {
      if (err instanceof DocxSerializeError) throw err;
      throw new DocxSerializeError("content-types-failed", "Failed to update [Content_Types].xml", {
        cause: err,
      });
    }
  }

  return container.serialize();
}

/**
 * Ensure `[Content_Types].xml` has a `<Default Extension="…">` entry for
 * every media part currently in the snapshot. We only add — never remove
 * — defaults, because removing one could break parts owned by other
 * workstreams that we don't model. Idempotent (no-op on second call).
 */
function ensureMediaContentTypes(container: ooxml.OoxmlContainer, snapshot: DocxSnapshot): void {
  const ct = ooxml.ContentTypes.load(container);
  let changed = false;
  for (const part of snapshot.root.media.values()) {
    const ext = extensionOf(part.partPath).toLowerCase();
    if (!ext) continue;
    if (!ct.hasDefault(ext)) {
      ct.addDefault(ext, part.mimeType);
      changed = true;
    }
  }
  if (changed) ct.writeBack(container);
}

function extensionOf(partPath: string): string {
  const dot = partPath.lastIndexOf(".");
  if (dot < 0) return "";
  return partPath.slice(dot + 1);
}

function serializeDocumentXml(doc: DocxDocument): string {
  const bodyChildren: unknown[] = serializeBodyBlocks(doc.body);
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
      // Per-table dirty marker: when `raw` is set, the table has not been
      // touched since parse and we re-emit the cached subtree byte-for-byte
      // (this is what preserves SHA-256 identity of `word/document.xml` when
      // a sibling table is mutated). When `raw` is absent, the table was
      // produced by a mutating command and must be regenerated from typed
      // fields.
      if (block.raw) return serializeTableFromRaw(block.raw);
      return serializeTable(block, serializeBlock);
    case "section-break":
      // P3.2: byte-preservation fast path. `raw` is dropped by mutating
      // commands (none in P3.2 itself) so its presence is the signal
      // that this section has not been touched since parse.
      if (block.raw) return opaqueToEntry(block.raw);
      return serializeSectionProperties(block.properties);
    case "opaque-block":
      // P2.3: when a content-wrapper carrier (SDT / fldSimple / mc:* /
      // smartTag / customXml) was unwrapped at parse time and a mutation
      // later flipped `subtreeDirty`, splice the typed `children` back into
      // the carrier's content slot. Otherwise re-emit the cached subtree
      // verbatim — that is what preserves byte-identical round-trip for
      // documents whose SDTs are read but never edited.
      if (block.subtreeDirty === true && block.children) {
        return reemitOpaqueBlockWithChildren(block.raw, block.children);
      }
      return opaqueToEntry(block.raw);
    case "wrapper-marker":
      // Wrapper markers are never serialized standalone — they must
      // appear in matched begin / end pairs that `serializeBodyBlocks`
      // consumes structurally. Hitting this branch means the body
      // model contains an orphan marker (or some other code path tried
      // to serialize a body block in isolation), which is a model-shape
      // bug we want to surface loudly rather than silently emit
      // malformed XML.
      throw new DocxSerializeError(
        "orphan-wrapper-marker",
        `Wrapper marker (${block.side}, wrapperId=${block.wrapperId}) cannot be serialized in isolation; ` +
          `it must be paired and processed via serializeBodyBlocks.`
      );
    default: {
      const _exhaustive: never = block;
      throw new DocxSerializeError("unknown-block", `Unknown block kind: ${JSON.stringify(_exhaustive)}`);
    }
  }
}

/**
 * Walk the document body and serialize every block, re-bracketing
 * lifted content-wrapper carriers (`<w:sdt>`, `mc:AlternateContent`,
 * `<w:fldSimple>`, `<w:smartTag>`, `<w:customXml>`) using their
 * `WrapperMarker` begin / end pairs.
 *
 * The parser inlines a wrapper's inner blocks into `body` so the page
 * chunker and renderer flow them as regular paragraphs (the TOC SDT,
 * for instance, used to be one giant atom block that orphaned its
 * heading on a page of its own). On serialize we have to reconstruct
 * the carrier envelope verbatim — `wrapperRaw` carries the original
 * subtree, and `rewriteContentSlot` swaps in the freshly serialized
 * inner blocks at exactly the slot they came from.
 *
 * Nested wrappers (an SDT inside an SDT, for example) work via a
 * simple stack: every begin pushes a frame whose inner blocks
 * accumulate as we walk; the matching end pops the frame, rebuilds
 * the carrier, and appends the result to its parent frame (or the
 * body root if the stack is now empty).
 *
 * Unmatched markers are a model-shape bug and throw — the parser's
 * `liftOpaqueBlock` always emits balanced pairs, so encountering an
 * orphan means a mutation produced a malformed body.
 */
interface WrapperFrame {
  readonly raw: OpaqueXml;
  readonly wrapperId: string;
  readonly inner: unknown[];
}

function serializeBodyBlocks(blocks: ReadonlyArray<BlockNode>): unknown[] {
  const root: unknown[] = [];
  const stack: WrapperFrame[] = [];

  const pushSerialized = (entry: unknown): void => {
    if (stack.length === 0) {
      root.push(entry);
    } else {
      stack[stack.length - 1].inner.push(entry);
    }
  };

  for (const block of blocks) {
    if (block.kind === "wrapper-marker") {
      if (block.side === "begin") {
        stack.push({ raw: block.wrapperRaw, wrapperId: block.wrapperId, inner: [] });
        continue;
      }
      // side === "end"
      const top = stack.pop();
      if (!top) {
        throw new DocxSerializeError(
          "orphan-wrapper-marker",
          `Wrapper end marker without matching begin (wrapperId=${block.wrapperId}).`
        );
      }
      if (top.wrapperId !== block.wrapperId) {
        throw new DocxSerializeError(
          "mismatched-wrapper-marker",
          `Wrapper end marker (wrapperId=${block.wrapperId}) does not match top of stack (wrapperId=${top.wrapperId}).`
        );
      }
      pushSerialized(rewriteContentSlot(top.raw, top.inner));
      continue;
    }
    pushSerialized(serializeBlock(block));
  }

  if (stack.length > 0) {
    const open = stack[stack.length - 1];
    throw new DocxSerializeError(
      "unclosed-wrapper-marker",
      `Body ended with ${stack.length} unclosed wrapper(s); top wrapperId=${open.wrapperId}.`
    );
  }

  return root;
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
    // The typed `<w:numPr>` is only emitted when no opaque carrier
    // exists for it (see check below). Untouched paragraphs round-trip
    // through the opaqueProps blob byte-equivalent; paragraphs touched
    // by `docx:set-paragraph-list` / `docx:remove-paragraph-list`
    // strip that blob so this typed branch kicks in.
    const carryOpaqueNumPr = props.opaqueProps?.some((o) => o.tag === "w:numPr") === true;
    if (!carryOpaqueNumPr) {
      out.push({
        "w:numPr": [
          { "w:ilvl": [], ":@": { "@_w:val": String(props.numbering.ilvl) } },
          { "w:numId": [], ":@": { "@_w:val": String(props.numbering.numId) } },
        ],
      });
    }
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
      return serializeRunOrFieldWrapper(node);
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
      if (node.subtreeDirty === true && node.children) {
        return reemitOpaqueInlineWithChildren(node.raw, node.children);
      }
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

/**
 * P3.4 / W15 — runs that contain exactly one {@link PageNumberFieldLeaf}
 * round-trip as `<w:fldSimple w:instr="…"><w:r>…</w:r></w:fldSimple>`
 * because that is the OOXML idiom Word writes (and the form the
 * parser detects). Mixed runs (text + field) fall back to the plain
 * `<w:r>` serialization — those don't exist when the leaf was
 * produced via `docx:insert-page-number`, but a future
 * AI-stitched mutation could create one and we'd rather degrade
 * gracefully than throw.
 */
function serializeRunOrFieldWrapper(r: Run): unknown {
  if (r.children.length !== 1) return serializeRun(r);
  const only = r.children[0];
  if (only.kind !== "page-number-field") return serializeRun(r);

  const innerRunChildren: unknown[] = [];
  const rPr = serializeRunProperties(r.properties);
  if (rPr) innerRunChildren.push(rPr);
  // Word always writes an inner `<w:t>` with the cached display
  // value. Use the captured cachedText when round-tripping a parsed
  // field, otherwise emit the field name as a sentinel placeholder
  // (Word will recompute on open).
  const display = only.cachedText ?? "#";
  innerRunChildren.push({
    "w:t": [{ "#text": display }],
    ":@": { "@_xml:space": "preserve" },
  });
  return {
    "w:fldSimple": [{ "w:r": innerRunChildren }],
    ":@": { "@_w:instr": only.instr },
  };
}

function serializeRunProperties(props: RunProperties): unknown | null {
  const out: unknown[] = [];
  const rFontsAttrs = buildRFontsAttrs(props);
  if (rFontsAttrs) {
    out.push(makeEl("w:rFonts", rFontsAttrs));
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
      // Skip opaque rFonts when we emitted a typed one. Any of the
      // typed font slots being set is enough to drop the captured
      // duplicate; otherwise the parser already round-tripped the
      // raw element verbatim.
      if (o.tag === "w:rFonts" && rFontsAttrs) continue;
      out.push(opaqueToEntry(o));
    }
  }
  if (out.length === 0) return null;
  return { "w:rPr": out };
}

/**
 * Build the `<w:rFonts>` attribute map from the typed multi-script
 * font slots on `props`. Returns `null` when no typed slot has been
 * set so the serializer can fall back to round-tripping the opaque
 * `w:rFonts` element via `opaqueProps`.
 *
 * `fontFamily` is mirrored to `w:ascii` AND `w:hAnsi` only when
 * `fontFamilyHAnsi` is unset — matches Word's behaviour for
 * single-script edits while still honouring an explicit hAnsi slot
 * coming from the source document.
 */
function buildRFontsAttrs(props: RunProperties): Record<string, string> | null {
  const attrs: Record<string, string> = {};
  if (props.fontFamily) attrs["w:ascii"] = props.fontFamily;
  if (props.fontFamilyHAnsi) {
    attrs["w:hAnsi"] = props.fontFamilyHAnsi;
  } else if (props.fontFamily) {
    attrs["w:hAnsi"] = props.fontFamily;
  }
  if (props.fontFamilyEastAsia) attrs["w:eastAsia"] = props.fontFamilyEastAsia;
  if (props.fontFamilyComplexScript) attrs["w:cs"] = props.fontFamilyComplexScript;
  if (props.fontFamilyAsciiTheme) attrs["w:asciiTheme"] = props.fontFamilyAsciiTheme;
  if (props.fontFamilyHAnsiTheme) attrs["w:hAnsiTheme"] = props.fontFamilyHAnsiTheme;
  if (props.fontFamilyEastAsiaTheme) attrs["w:eastAsiaTheme"] = props.fontFamilyEastAsiaTheme;
  if (props.fontFamilyComplexScriptTheme) attrs["w:cstheme"] = props.fontFamilyComplexScriptTheme;
  return Object.keys(attrs).length === 0 ? null : attrs;
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
    case "page-break":
      return makeEl("w:br", { "w:type": "page" });
    case "last-rendered-page-break":
      return { "w:lastRenderedPageBreak": [] };
    case "page-number-field": {
      // Defensive path: page-number leaves should be lifted to a
      // `<w:fldSimple>` wrapper at the inline level by
      // `serializeRunOrFieldWrapper`. If the leaf reaches here it
      // means a *mixed* run snuck through; emit a `<w:fldSimple>`
      // anyway so the field semantics survive, even though the
      // surrounding text in the same run will end up grouped under
      // a sibling `<w:r>` by the caller.
      return {
        "w:fldSimple": [
          { "w:r": [{ "w:t": [{ "#text": c.cachedText ?? "#" }], ":@": { "@_xml:space": "preserve" } }] },
        ],
        ":@": { "@_w:instr": c.instr },
      };
    }
    case "tab":
      return { "w:tab": [] };
    case "drawing": {
      switch (c.subkind) {
        case "inline-image":
          // Byte-preservation fast path: re-emit cached subtree when the
          // typed model has not been touched since parse.
          if (c.raw) return opaqueToEntry(c.raw);
          return serializeInlineImageDrawing(c);
        case "chart":
          // The chart payload itself lives in `word/charts/chartN.xml`
          // and is rewritten by the dirty-charts pass below; the inline
          // drawing wrapper round-trips via the captured `raw` subtree
          // for parsed charts and via a synthesized envelope for newly
          // inserted charts.
          if (c.raw) return opaqueToEntry(c.raw);
          return serializeChartDrawing(c);
        case "opaque":
          return opaqueToEntry(c.raw);
        default: {
          const _exhaustive: never = c;
          throw new DocxSerializeError(
            "unknown-drawing-subkind",
            `Unknown drawing subkind: ${JSON.stringify(_exhaustive)}`
          );
        }
      }
    }
    case "embedded-spreadsheet":
      return serializeEmbeddedSpreadsheet(c);
    case "footnote-ref": {
      // F1 — promoted typed leaf round-trips back to
      // `<w:footnoteReference w:id="N"/>`. We don't carry a `raw`
      // envelope here because the element has zero unmodelled state
      // beyond its two attributes, so re-emission is exact.
      const attrs: Record<string, string> = { "w:id": String(c.footnoteId) };
      if (c.customMarkFollows) attrs["w:customMarkFollows"] = "1";
      return makeEl("w:footnoteReference", attrs);
    }
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

/**
 * Re-emit `<w:object>` for a typed OLE Excel embed. Existing embeds
 * round-trip via the captured `raw` subtree (same byte-preservation
 * fast path images use). Authored embeds with no captured raw fall
 * back to a minimal `<w:object>` envelope carrying the rebuilt
 * `<o:OLEObject>` and a `<v:shape>` preview wrapper. The full
 * `<v:imagedata>` / preview-image authoring lives in the serializer
 * companion that runs alongside `docx:insert-spreadsheet`.
 */
function serializeEmbeddedSpreadsheet(leaf: import("../model/types.js").EmbeddedSpreadsheet): unknown {
  if (leaf.raw) return opaqueToEntry(leaf.raw);
  const oleAttrs: Record<string, string> = { ...leaf.oleObjectAttrs };
  oleAttrs["ProgID"] = leaf.progId;
  oleAttrs["r:id"] = leaf.oleRelId;
  if (!oleAttrs["Type"]) oleAttrs["Type"] = "Embed";
  if (!oleAttrs["DrawAspect"]) oleAttrs["DrawAspect"] = "Content";
  const oleObj = makeEntry("o:OLEObject", [], oleAttrs);
  const objectChildren: unknown[] = [];
  if (leaf.previewImageRelId) {
    const shapeAttrs: Record<string, string> = {
      id: `_x0000_i${leaf.id}`,
      type: "#_x0000_t75",
      style: "width:240pt;height:180pt",
    };
    const imagedata = makeEntry("v:imagedata", [], { "r:id": leaf.previewImageRelId, "o:title": "" });
    const shape = makeEntry("v:shape", [imagedata], shapeAttrs);
    objectChildren.push(shape);
  }
  objectChildren.push(oleObj);
  return makeEntry("w:object", objectChildren, {});
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

/**
 * Serialize an `OpaqueBlock` whose `subtreeDirty` flag was flipped by a
 * mutation: walk the cached `raw` subtree, locate the wrapper's content
 * slot (e.g. `<w:sdtContent>` for `<w:sdt>`), and replace its element
 * children with freshly-serialized typed `BlockNode` children. Non-element
 * markup (comments, processing instructions, attributes) inside the slot
 * is preserved.
 *
 * Tags whose direct children are the content slot (e.g. `<w:fldSimple>`,
 * `<w:smartTag>`) get their own children replaced. For wrappers that
 * carry siblings adjacent to the content slot (`<w:sdt>` has `<w:sdtPr>`
 * + `<w:sdtEndPr>` + `<w:sdtContent>`), only the slot is rewritten.
 */
function reemitOpaqueBlockWithChildren(raw: OpaqueXml, children: ReadonlyArray<BlockNode>): unknown {
  const newChildren: unknown[] = children.map((c) => serializeBlock(c));
  const rewritten = rewriteContentSlot(raw, newChildren);
  return rewritten;
}

function reemitOpaqueInlineWithChildren(raw: OpaqueXml, children: ReadonlyArray<InlineNode>): unknown {
  const newChildren: unknown[] = children.map((c) => serializeInline(c));
  return rewriteContentSlot(raw, newChildren);
}

function rewriteContentSlot(raw: OpaqueXml, newChildren: unknown[]): unknown {
  const tag = raw.tag;
  switch (tag) {
    case "w:sdt": {
      // Walk the subtree array, replace `<w:sdtContent>`'s inner children.
      const subtree = raw.subtree.map((entry) => {
        if (!entry || typeof entry !== "object" || Array.isArray(entry)) return entry;
        const obj = entry as Record<string, unknown>;
        const keys = Object.keys(obj).filter((k) => k !== ATTR_KEY);
        if (keys.length === 1 && keys[0] === "w:sdtContent") {
          const next: Record<string, unknown> = { "w:sdtContent": newChildren };
          if (obj[ATTR_KEY]) next[ATTR_KEY] = obj[ATTR_KEY];
          return next;
        }
        return entry;
      });
      const out: Record<string, unknown> = { [tag]: subtree };
      if (Object.keys(raw.rawAttrs).length > 0) out[ATTR_KEY] = { ...raw.rawAttrs };
      return out;
    }
    case "mc:AlternateContent": {
      // Rewrite the first <mc:Choice> (or <mc:Fallback>) we find.
      let replaced = false;
      const subtree = raw.subtree.map((entry) => {
        if (!entry || typeof entry !== "object" || Array.isArray(entry)) return entry;
        const obj = entry as Record<string, unknown>;
        const keys = Object.keys(obj).filter((k) => k !== ATTR_KEY);
        if (keys.length !== 1) return entry;
        const childTag = keys[0];
        if (replaced) return entry;
        if (childTag === "mc:Choice" || childTag === "mc:Fallback") {
          replaced = true;
          const next: Record<string, unknown> = { [childTag]: newChildren };
          if (obj[ATTR_KEY]) next[ATTR_KEY] = obj[ATTR_KEY];
          return next;
        }
        return entry;
      });
      const out: Record<string, unknown> = { [tag]: subtree };
      if (Object.keys(raw.rawAttrs).length > 0) out[ATTR_KEY] = { ...raw.rawAttrs };
      return out;
    }
    case "w:sdtContent":
    case "mc:Choice":
    case "mc:Fallback":
    case "w:fldSimple":
    case "w:smartTag":
    case "w:customXml": {
      const out: Record<string, unknown> = { [tag]: newChildren };
      if (Object.keys(raw.rawAttrs).length > 0) out[ATTR_KEY] = { ...raw.rawAttrs };
      return out;
    }
    default:
      // No known content slot for this tag — fall back to the cached subtree
      // verbatim. This shouldn't happen in practice because the parser only
      // attaches `children` when `blockContentSlot` returned non-null.
      return opaqueToEntry(raw);
  }
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

/**
 * B7 — register `word/numbering.xml` with the package when the typed
 * `NumberingDefinitions` carrier becomes dirty for the first time.
 *
 * The serializer for `word/numbering.xml` writes the bytes; this
 * helper makes sure the surrounding plumbing exists so Word actually
 * picks the part up:
 *
 *   1. The `word/_rels/document.xml.rels` graph carries a relationship
 *      of type `…/relationships/numbering` pointing at the part.
 *   2. `[Content_Types].xml` declares an `<Override>` mapping the
 *      part path to the canonical numbering content-type.
 *
 * Idempotent on every call: existing relationships / overrides are
 * left alone, so re-saving a doc that already had `word/numbering.xml`
 * stays byte-identical for the surrounding parts.
 */
function ensureNumberingPart(container: ooxml.OoxmlContainer): void {
  const rels = ooxml.RelationshipGraph.loadFor(container, MAIN_PART);
  if (rels.byType(NUMBERING_REL_TYPE).length === 0) {
    rels.add({ type: NUMBERING_REL_TYPE, target: "numbering.xml" });
    rels.writeBack(container);
  }
  const ct = ooxml.ContentTypes.load(container);
  if (!ct.hasOverride("/word/numbering.xml")) {
    ct.addOverride("/word/numbering.xml", NUMBERING_CONTENT_TYPE);
    ct.writeBack(container);
  }
  void NUMBERING_PART;
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
