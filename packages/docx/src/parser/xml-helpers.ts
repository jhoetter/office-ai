import { ooxml } from "@officeai/core";
import type { OpaqueXml } from "../model/types.js";

export const ATTR_KEY = ":@";
export const ATTR_PREFIX = ooxml.xmlAttrPrefix;

/**
 * Capture an element entry as an OpaqueXml carrier so the serializer can
 * re-emit it byte-equivalent.
 */
export function captureOpaque(entry: Record<string, unknown>): OpaqueXml {
  const tag = ooxml.getTag(entry);
  const attrs = ooxml.getAttrs(entry);
  const subtree = (entry[tag] as unknown[] | undefined) ?? [];
  const rawAttrs: Record<string, string> = {};
  const attrMap = entry[ATTR_KEY];
  if (attrMap && typeof attrMap === "object") {
    for (const [k, v] of Object.entries(attrMap as Record<string, unknown>)) {
      rawAttrs[k] = String(v);
    }
  }
  return { tag, attrs, subtree, rawAttrs };
}

/**
 * Re-emit an OpaqueXml as a fast-xml-parser preserveOrder element entry.
 */
export function opaqueToEntry(o: OpaqueXml): Record<string, unknown> {
  const entry: Record<string, unknown> = { [o.tag]: o.subtree };
  if (Object.keys(o.rawAttrs).length > 0) {
    entry[ATTR_KEY] = { ...o.rawAttrs };
  }
  return entry;
}

/**
 * Walk the top-level array, extracting only element entries (skips comments,
 * the `?xml` declaration, etc.).
 */
export function elementEntries(siblings: ReadonlyArray<unknown>): Array<Record<string, unknown>> {
  const out: Array<Record<string, unknown>> = [];
  for (const s of siblings) {
    if (!s || typeof s !== "object" || Array.isArray(s)) continue;
    const obj = s as Record<string, unknown>;
    const keys = Object.keys(obj).filter((k) => k !== ATTR_KEY);
    if (keys.length === 1) {
      const tag = keys[0];
      if (!tag.startsWith("?") && tag !== "#text" && tag !== "#comment") {
        out.push(obj);
      }
    }
  }
  return out;
}

export function findElementEntry(
  siblings: ReadonlyArray<unknown>,
  tag: string,
): Record<string, unknown> | null {
  for (const e of elementEntries(siblings)) {
    if (ooxml.getTag(e) === tag) return e;
  }
  return null;
}

export function rootEntry(tree: unknown, expectedTag: string): Record<string, unknown> {
  if (!Array.isArray(tree)) {
    throw new Error(`expected XML root array; got ${typeof tree}`);
  }
  const found = findElementEntry(tree as unknown[], expectedTag);
  if (!found) {
    throw new Error(`root element <${expectedTag}> not found`);
  }
  return found;
}

/**
 * Read text content of an element by concatenating all `#text` children.
 */
export function readText(entry: Record<string, unknown>): string {
  const tag = ooxml.getTag(entry);
  const children = entry[tag];
  if (!Array.isArray(children)) return "";
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

export function attrOf(entry: Record<string, unknown>, name: string): string | undefined {
  const a = entry[ATTR_KEY];
  if (!a || typeof a !== "object") return undefined;
  const map = a as Record<string, unknown>;
  const v = map[`${ATTR_PREFIX}${name}`];
  return v === undefined ? undefined : String(v);
}
