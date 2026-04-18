import { ooxml } from "@officeai/core";
import type { ParagraphProperties, RunProperties, StyleDefinition, StylesPart } from "../model/types.js";
import { DocxParseError } from "./errors.js";
import { parseParagraphProperties, parseRunProperties } from "./parse.js";
import { attrOf, captureOpaque, elementEntries, findElementEntry } from "./xml-helpers.js";

const STYLES_PART = "word/styles.xml";

/**
 * Parse `word/styles.xml` into a typed `StylesPart`. Returns `undefined`
 * when the part is absent (synthetic test fixtures often omit it).
 *
 * Byte-preservation: the typed projection is read-only metadata for
 * the cascade resolver. The serializer never writes back from it
 * unless `dirty.styles` is set, which only happens when a future
 * workstream mutates a style definition (P4 / R10). The original
 * `<w:styles>` subtree is captured into `raw` so a hypothetical
 * "render from typed model" code path can re-emit byte-identical
 * output for unmodified docs.
 */
export function parseStylesPart(container: ooxml.OoxmlContainer): StylesPart | undefined {
  if (!container.has(STYLES_PART)) return undefined;

  let tree: unknown;
  try {
    tree = ooxml.parseXml(container.readText(STYLES_PART));
  } catch (err) {
    throw new DocxParseError("invalid-xml", "Failed to parse styles.xml", {
      partPath: STYLES_PART,
      cause: err,
    });
  }
  if (!Array.isArray(tree)) {
    return { docDefaults: {}, styles: new Map<string, StyleDefinition>() };
  }
  const root = findElementEntry(tree as unknown[], "w:styles");
  if (!root) {
    return { docDefaults: {}, styles: new Map<string, StyleDefinition>() };
  }

  const styles = new Map<string, StyleDefinition>();
  let docDefaults: StylesPart["docDefaults"] = {};

  for (const entry of elementEntries((root["w:styles"] as unknown[] | undefined) ?? [])) {
    const tag = ooxml.getTag(entry);
    if (tag === "w:docDefaults") {
      docDefaults = parseDocDefaults(entry);
    } else if (tag === "w:style") {
      const def = parseStyleDefinition(entry);
      if (def) styles.set(def.id, def);
    }
  }

  return { docDefaults, styles, raw: captureOpaque(root) };
}

function parseDocDefaults(entry: Record<string, unknown>): StylesPart["docDefaults"] {
  const children = (entry["w:docDefaults"] as unknown[] | undefined) ?? [];
  let rPrDefault: RunProperties | undefined;
  let pPrDefault: ParagraphProperties | undefined;

  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    if (tag === "w:rPrDefault") {
      const inner = findElementEntry((c["w:rPrDefault"] as unknown[] | undefined) ?? [], "w:rPr");
      if (inner) rPrDefault = parseRunProperties(inner);
    } else if (tag === "w:pPrDefault") {
      const inner = findElementEntry((c["w:pPrDefault"] as unknown[] | undefined) ?? [], "w:pPr");
      if (inner) pPrDefault = parseParagraphProperties(inner);
    }
  }

  const out: StylesPart["docDefaults"] = {};
  if (rPrDefault) (out as { rPrDefault?: RunProperties }).rPrDefault = rPrDefault;
  if (pPrDefault) (out as { pPrDefault?: ParagraphProperties }).pPrDefault = pPrDefault;
  return out;
}

function parseStyleDefinition(entry: Record<string, unknown>): StyleDefinition | null {
  const id = attrOf(entry, "w:styleId");
  if (!id) return null;
  const typeAttr = attrOf(entry, "w:type");
  const type = normalizeStyleType(typeAttr);
  if (!type) return null;

  const children = (entry["w:style"] as unknown[] | undefined) ?? [];
  const def: { -readonly [K in keyof StyleDefinition]: StyleDefinition[K] } = {
    id,
    type,
  };
  const opaqueProps: ReturnType<typeof captureOpaque>[] = [];

  if (attrOf(entry, "w:default") === "1") def.default = true;

  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    switch (tag) {
      case "w:name": {
        const v = attrOf(c, "w:val");
        if (v) def.name = v;
        break;
      }
      case "w:basedOn": {
        const v = attrOf(c, "w:val");
        if (v) def.basedOn = v;
        break;
      }
      case "w:next": {
        const v = attrOf(c, "w:val");
        if (v) def.next = v;
        break;
      }
      case "w:link": {
        const v = attrOf(c, "w:val");
        if (v) def.link = v;
        break;
      }
      case "w:hidden":
        def.hidden = true;
        break;
      case "w:rPr":
        def.rPr = parseRunProperties(c);
        break;
      case "w:pPr":
        def.pPr = parseParagraphProperties(c);
        break;
      default:
        opaqueProps.push(captureOpaque(c));
        break;
    }
  }

  if (opaqueProps.length > 0) def.opaqueProps = opaqueProps;
  return def;
}

function normalizeStyleType(v: string | undefined): StyleDefinition["type"] | null {
  switch (v) {
    case "paragraph":
    case "character":
    case "table":
    case "numbering":
      return v;
    default:
      return null;
  }
}
