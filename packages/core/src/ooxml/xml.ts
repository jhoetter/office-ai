import { XMLBuilder, XMLParser } from "fast-xml-parser";

/**
 * fast-xml-parser preserveOrder shape:
 *   XmlNode = Array<{ "<tag>": XmlNode | string; ":@"?: Record<string,string>; "#text"?: string }>
 *
 * We type it loosely as `unknown` at the public boundary; helpers below
 * narrow as needed.
 */
export type XmlNode = unknown;

export const xmlAttrPrefix = "@_";

const PARSER_OPTIONS = {
  ignoreAttributes: false,
  attributeNamePrefix: xmlAttrPrefix,
  allowBooleanAttributes: true,
  parseTagValue: false,
  parseAttributeValue: false,
  preserveOrder: true,
  trimValues: false,
  unpairedTags: ["w:br", "w:tab", "w:cr"] as string[],
};

const BUILDER_OPTIONS = {
  ignoreAttributes: false,
  attributeNamePrefix: xmlAttrPrefix,
  allowBooleanAttributes: true,
  preserveOrder: true,
  format: false,
  suppressEmptyNode: false,
  suppressBooleanAttributes: false,
  unpairedTags: ["w:br", "w:tab", "w:cr"] as string[],
};

const parser = new XMLParser(PARSER_OPTIONS);
const builder = new XMLBuilder(BUILDER_OPTIONS);

export function parseXml(xml: string): XmlNode {
  return parser.parse(xml) as XmlNode;
}

/**
 * Serialize. By default we emit `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
 * unless the tree itself contains a `?xml` declaration (parser will surface it).
 */
export function serializeXml(tree: XmlNode, opts?: { xmlDeclaration?: string | null }): string {
  const xml = builder.build(tree);
  const decl = opts?.xmlDeclaration === undefined
    ? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    : opts.xmlDeclaration === null
      ? ""
      : opts.xmlDeclaration;
  if (typeof xml !== "string") return decl;
  if (xml.startsWith("<?xml")) return xml;
  return decl + xml;
}

/* ── Helpers for navigating the preserveOrder tree ───────────────────────── */

export interface XmlElement {
  /** The single element entry, e.g. { "w:p": [...children], ":@": {attrs} }. */
  readonly entry: Record<string, unknown>;
  readonly tag: string;
  readonly children: ReadonlyArray<unknown>;
  readonly attrs: Readonly<Record<string, string>>;
}

const ATTR_KEY = ":@";

export function isElementEntry(node: unknown): node is Record<string, unknown> {
  if (!node || typeof node !== "object" || Array.isArray(node)) return false;
  const keys = Object.keys(node);
  // An element entry has exactly one tag key (plus optional ":@").
  const tagKeys = keys.filter((k) => k !== ATTR_KEY);
  return tagKeys.length === 1;
}

export function getTag(entry: Record<string, unknown>): string {
  for (const k of Object.keys(entry)) {
    if (k !== ATTR_KEY) return k;
  }
  return "";
}

export function getChildren(entry: Record<string, unknown>): ReadonlyArray<unknown> {
  const tag = getTag(entry);
  const v = entry[tag];
  if (Array.isArray(v)) return v as unknown[];
  return [];
}

export function getAttrs(entry: Record<string, unknown>): Readonly<Record<string, string>> {
  const a = entry[ATTR_KEY];
  if (!a || typeof a !== "object") return {};
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(a as Record<string, unknown>)) {
    if (k.startsWith(xmlAttrPrefix)) {
      out[k.slice(xmlAttrPrefix.length)] = String(v);
    } else {
      out[k] = String(v);
    }
  }
  return out;
}

export function asElement(node: unknown): XmlElement | null {
  if (!isElementEntry(node)) return null;
  return {
    entry: node,
    tag: getTag(node),
    children: getChildren(node),
    attrs: getAttrs(node),
  };
}

export function findChild(siblings: ReadonlyArray<unknown>, tag: string): XmlElement | null {
  for (const s of siblings) {
    const el = asElement(s);
    if (el && el.tag === tag) return el;
  }
  return null;
}

export function filterChildren(siblings: ReadonlyArray<unknown>, tag: string): XmlElement[] {
  const out: XmlElement[] = [];
  for (const s of siblings) {
    const el = asElement(s);
    if (el && el.tag === tag) out.push(el);
  }
  return out;
}

/** Return the text content of a `<w:t>` (or any element whose only child is `#text`). */
export function getTextContent(entry: Record<string, unknown>): string {
  const children = getChildren(entry);
  let out = "";
  for (const c of children) {
    if (c && typeof c === "object" && !Array.isArray(c)) {
      const obj = c as Record<string, unknown>;
      const t = obj["#text"];
      if (t !== undefined) out += String(t);
    }
  }
  return out;
}

/* ── Builders for new entries ────────────────────────────────────────────── */

export interface AttrMap {
  [key: string]: string | number | boolean | undefined;
}

/** Build a fast-xml-parser preserveOrder element entry. */
export function makeElement(
  tag: string,
  children: ReadonlyArray<unknown> = [],
  attrs?: AttrMap,
): Record<string, unknown> {
  const entry: Record<string, unknown> = { [tag]: children };
  if (attrs) {
    const out: Record<string, string> = {};
    for (const [k, v] of Object.entries(attrs)) {
      if (v === undefined) continue;
      out[`${xmlAttrPrefix}${k}`] = String(v);
    }
    if (Object.keys(out).length > 0) entry[ATTR_KEY] = out;
  }
  return entry;
}

/** Build a `#text` leaf. */
export function makeTextLeaf(text: string): Record<string, unknown> {
  return { "#text": text };
}
