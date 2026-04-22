import { ooxml, type IdMinter } from "@officeai/core";
import type { BlockNode, Footnote, FootnoteType, FootnotesPart, OpaqueXml } from "../model/types.js";
import { DocxParseError } from "./errors.js";
import { attrOf, captureOpaque, elementEntries, findElementEntry } from "./xml-helpers.js";

/**
 * Parser for `word/footnotes.xml` (F1).
 *
 * Returns `undefined` when the part is absent — the common case for
 * letters, short reports, and any document the user has never inserted
 * a footnote into. Returns a populated `FootnotesPart` for documents
 * that carry the standard `separator` / `continuationSeparator` notes
 * Word inserts on first footnote authoring, plus any user `normal`
 * footnotes.
 *
 * Round-trip contract: every parsed `<w:footnote>` keeps its full
 * subtree on `Footnote.raw`, so the serializer re-emits unchanged
 * footnotes byte-for-byte. Mutating commands drop `raw` on the
 * footnote they touch; the serializer then regenerates that one
 * footnote from the typed model while leaving siblings alone.
 *
 * `parseParagraph` is injected (not imported) to avoid an import cycle
 * with `parse.ts`; same pattern as `parser/headers-footers.ts`.
 */
const FOOTNOTES_PART = "word/footnotes.xml";
const FOOTNOTES_TAG = "w:footnotes";
const FOOTNOTE_TAG = "w:footnote";

const ATTR_KEY = ":@";

export type ParseFootnoteParagraph = (entry: Record<string, unknown>, mintNodeId: IdMinter) => BlockNode;

export type ParseFootnoteTable = (entry: Record<string, unknown>, mintNodeId: IdMinter) => BlockNode;

export function parseFootnotesPart(
  container: ooxml.OoxmlContainer,
  mintNodeId: IdMinter,
  parseParagraph: ParseFootnoteParagraph,
  parseTable?: ParseFootnoteTable
): FootnotesPart | undefined {
  if (!container.has(FOOTNOTES_PART)) return undefined;

  const xml = container.readText(FOOTNOTES_PART);
  let tree: unknown;
  try {
    tree = ooxml.parseXml(xml);
  } catch (err) {
    throw new DocxParseError("invalid-xml", `Failed to parse ${FOOTNOTES_PART}`, {
      partPath: FOOTNOTES_PART,
      cause: err,
    });
  }
  if (!Array.isArray(tree)) {
    return { footnotes: [], rootAttrs: {} };
  }
  const rootEntry = findElementEntry(tree as unknown[], FOOTNOTES_TAG);
  if (!rootEntry) {
    return { footnotes: [], rootAttrs: {} };
  }

  const rootAttrs = readRootAttrs(rootEntry);
  const children = (rootEntry[FOOTNOTES_TAG] as unknown[] | undefined) ?? [];
  const footnotes: Footnote[] = [];
  const tail: OpaqueXml[] = [];

  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    if (tag !== FOOTNOTE_TAG) {
      tail.push(captureOpaque(c));
      continue;
    }

    const idAttr = attrOf(c, "w:id");
    if (idAttr === undefined) {
      // Malformed: a `<w:footnote>` without `w:id` is ill-formed
      // OOXML. Round-trip it as opaque tail so we don't lose bytes
      // but don't surface it in the typed list either.
      tail.push(captureOpaque(c));
      continue;
    }
    const id = Number(idAttr);
    if (!Number.isFinite(id)) {
      tail.push(captureOpaque(c));
      continue;
    }

    const type = normalizeFootnoteType(attrOf(c, "w:type"));
    const body: BlockNode[] = [];
    const innerChildren = (c[FOOTNOTE_TAG] as unknown[] | undefined) ?? [];
    for (const inner of elementEntries(innerChildren)) {
      const innerTag = ooxml.getTag(inner);
      if (innerTag === "w:p") {
        body.push(parseParagraph(inner, mintNodeId));
      } else if (innerTag === "w:tbl" && parseTable) {
        body.push(parseTable(inner, mintNodeId));
      } else {
        body.push({ kind: "opaque-block", id: mintNodeId(), raw: captureOpaque(inner) });
      }
    }

    footnotes.push({
      id,
      type,
      body,
      raw: captureOpaque(c),
    });
  }

  const part: {
    -readonly [K in keyof FootnotesPart]: FootnotesPart[K];
  } = { footnotes, rootAttrs };
  if (tail.length > 0) part.tail = tail;
  return part;
}

function normalizeFootnoteType(t: string | undefined): FootnoteType {
  if (t === "separator" || t === "continuationSeparator" || t === "continuationNotice") return t;
  return "normal";
}

function readRootAttrs(entry: Record<string, unknown>): Record<string, string> {
  const attrs: Record<string, string> = {};
  const a = entry[ATTR_KEY];
  if (!a || typeof a !== "object") return attrs;
  for (const [k, v] of Object.entries(a as Record<string, unknown>)) {
    const name = k.startsWith("@_") ? k.slice(2) : k;
    attrs[name] = String(v);
  }
  return attrs;
}
