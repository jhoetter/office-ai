import { ooxml } from "@officeai/core";
import type { Comment } from "../model/types.js";
import { XlsxParseError } from "./errors.js";

/**
 * Parse `xl/comments{N}.xml` into a typed `{ authors, comments }` pair.
 *
 * Each `<comment>` element references an author by `authorId` index
 * into `<authors><author>…`. Rich-text runs (`<r><t>`) are flattened
 * to a single plain-text string for P0 — round-tripping through
 * `serializeCommentsPart` re-emits a single `<r><t>` run, which is
 * the documented P0 trade-off (rich-text formatting is not modeled).
 *
 * Comment ids are minted positionally as `comment-1`, `comment-2`, …
 * — stable per file, opaque to callers.
 */
export function parseCommentsPart(
  xml: string,
  partPath?: string
): { authors: string[]; comments: Comment[] } {
  let tree: unknown;
  try {
    tree = ooxml.parseXml(xml);
  } catch (err) {
    throw new XlsxParseError("invalid-xml", `Failed to parse comments part`, {
      partPath,
      cause: err,
    });
  }
  if (!Array.isArray(tree)) return { authors: [], comments: [] };

  const root = (tree as unknown[])
    .map((n) => ooxml.asElement(n))
    .find((el): el is ooxml.XmlElement => el !== null && el.tag === "comments");
  if (!root) return { authors: [], comments: [] };

  const authors: string[] = [];
  const authorsEl = ooxml.findChild(root.children, "authors");
  if (authorsEl) {
    for (const a of ooxml.filterChildren(authorsEl.children, "author")) {
      authors.push(ooxml.getTextContent(a.entry));
    }
  }

  const comments: Comment[] = [];
  const listEl = ooxml.findChild(root.children, "commentList");
  if (listEl) {
    let i = 0;
    for (const c of ooxml.filterChildren(listEl.children, "comment")) {
      i++;
      const ref = c.attrs.ref ?? "";
      const authorIdRaw = c.attrs.authorId;
      const authorIdx = authorIdRaw === undefined ? -1 : Number.parseInt(authorIdRaw, 10);
      const author =
        Number.isInteger(authorIdx) && authorIdx >= 0 && authorIdx < authors.length ? authors[authorIdx] : "";
      const text = extractCommentText(c);
      comments.push({ id: `comment-${i}`, ref, author, text });
    }
  }

  return { authors, comments };
}

function extractCommentText(commentEl: ooxml.XmlElement): string {
  const textEl = ooxml.findChild(commentEl.children, "text");
  if (!textEl) return "";

  let out = "";
  // Direct `<t>` children (rare but legal) plus `<r><t>` runs.
  for (const child of textEl.children) {
    const el = ooxml.asElement(child);
    if (!el) continue;
    if (el.tag === "t") {
      out += ooxml.getTextContent(el.entry);
      continue;
    }
    if (el.tag === "r") {
      const runT = ooxml.findChild(el.children, "t");
      if (runT) out += ooxml.getTextContent(runT.entry);
    }
  }
  return out;
}
