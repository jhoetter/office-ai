import type { OpaqueXml } from "./types.js";

/**
 * Display-side classification for opaque OOXML carriers.
 *
 * Real-world `.docx` files emit a meaningful number of "structural" elements
 * that are zero-width metadata (bookmarks, field characters, paragraph-level
 * proof markers) or lossless wrappers around editable content (SDT content
 * controls, simple fields, MC fallback content). The model preserves them all
 * byte-for-byte through `OpaqueXml`, but a naive renderer that surfaces each
 * one as `[opaque]` or `[<tag>]` clutters the editing surface to the point of
 * being unusable on any non-trivial document.
 *
 * `classifyOpaqueTag` is the single source of truth used by the renderer to
 * decide HOW to display an opaque carrier without changing what we round-trip.
 *
 *   - `metadata`        → emit nothing (the carrier is invisible markup).
 *   - `content-wrapper` → render the carrier's flattened inner text in place
 *                         of the wrapper, so the user sees readable content.
 *   - `placeholder`     → fall back to the legacy `[<tag>]` chip (default).
 */
export type OpaqueDisplay = "metadata" | "content-wrapper" | "placeholder";

const METADATA_TAGS: ReadonlySet<string> = new Set([
  "w:bookmarkStart",
  "w:bookmarkEnd",
  "w:proofErr",
  "w:lastRenderedPageBreak",
  "w:fldChar",
  "w:instrText",
  "w:delInstrText",
  "w:permStart",
  "w:permEnd",
  "w:moveFromRangeStart",
  "w:moveFromRangeEnd",
  "w:moveToRangeStart",
  "w:moveToRangeEnd",
  "w:customXmlInsRangeStart",
  "w:customXmlInsRangeEnd",
  "w:customXmlDelRangeStart",
  "w:customXmlDelRangeEnd",
  "w:customXmlMoveFromRangeStart",
  "w:customXmlMoveFromRangeEnd",
  "w:customXmlMoveToRangeStart",
  "w:customXmlMoveToRangeEnd",
  "w:annotationRef",
  "w:footnoteRef",
  "w:endnoteRef",
  "w:rsidRoot",
]);

const CONTENT_WRAPPER_TAGS: ReadonlySet<string> = new Set([
  "w:sdt",
  "w:sdtContent",
  "w:fldSimple",
  "mc:AlternateContent",
  "mc:Choice",
  "mc:Fallback",
  "w:smartTag",
  "w:customXml",
]);

export function classifyOpaqueTag(tag: string): OpaqueDisplay {
  if (METADATA_TAGS.has(tag)) return "metadata";
  if (CONTENT_WRAPPER_TAGS.has(tag)) return "content-wrapper";
  return "placeholder";
}

/**
 * Recursively flatten the visible text inside an `OpaqueXml` subtree by
 * concatenating every `#text` descendant. Returns the empty string for
 * subtrees with no text content.
 *
 * The walker is preserveOrder-aware: `subtree` is an array of element entries
 * (each entry is `{ "<tag>": children, ":@"?: attrs }` or `{ "#text": "..." }`).
 */
export function extractOpaqueText(opaque: OpaqueXml): string {
  return collectText(opaque.subtree);
}

function collectText(siblings: ReadonlyArray<unknown>): string {
  let out = "";
  for (const node of siblings) {
    if (!node || typeof node !== "object" || Array.isArray(node)) continue;
    const obj = node as Record<string, unknown>;
    if (typeof obj["#text"] === "string") {
      out += obj["#text"] as string;
      continue;
    }
    for (const [key, value] of Object.entries(obj)) {
      if (key === ":@" || key.startsWith("?")) continue;
      // Skip the textual content of nested metadata-only carriers (e.g. the
      // `<w:instrText>` field instruction inside a `<w:sdt>` TOC wrapper):
      // it is a runtime construct, not user-visible content, and surfacing
      // it as preview text would just leak XML noise back into the editor.
      if (classifyOpaqueTag(key) === "metadata") continue;
      if (Array.isArray(value)) {
        out += collectText(value);
      }
    }
  }
  return out;
}
