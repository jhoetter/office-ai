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
      // Capture the entire `<text>…</text>` element as a verbatim
      // string so the serializer can re-emit it when the user
      // hasn't actually edited the comment body. This preserves
      // rich-text runs (multiple `<r>` children with their own
      // `<rPr>` font/color/etc.), arbitrary whitespace, and any
      // unmodeled run children that the typed flattening would
      // otherwise discard. Extracted from the original XML by
      // pinpointing the `<text>` tag inside this comment's `<comment>`
      // element. We use a scoped regex (anchored on the comment's
      // ref attr) because we don't have direct access to the
      // serialized form of `c` from the `XmlElement` wrapper.
      const textXml = extractTextElementForComment(xml, ref, c.attrs.authorId);
      // P1: round-trip threaded-comment metadata that
      // `serializer/comments.ts` writes as `officeai-*` attributes.
      // Excel preserves unknown attributes, so this survives a real
      // Excel save → re-load too.
      const parentId = c.attrs["officeai-parentId"];
      const resolvedRaw = c.attrs["officeai-resolved"];
      const createdAt = c.attrs["officeai-createdAt"];
      comments.push({
        id: `comment-${i}`,
        ref,
        author,
        text,
        ...(parentId ? { parentId } : {}),
        ...(resolvedRaw === "1" ? { resolved: true } : {}),
        ...(createdAt ? { createdAt } : {}),
        ...(textXml ? { textXml } : {}),
      });
    }
  }

  return { authors, comments };
}

/**
 * Pull the `<text>…</text>` element verbatim for a single comment.
 *
 * Scoped via the comment's `ref` attribute (and `authorId` as a
 * tiebreaker when two comments share a ref). Returns `undefined`
 * when no match is found — the caller falls back to flattening
 * via {@link extractCommentText}, which is lossy but safe.
 *
 * NOTE: This is a regex-based scan so it tolerates the various
 * orderings of `ref`, `authorId`, and our `officeai-*` attrs that
 * different writers emit. The scoped open tag uses the actual
 * attribute order from the source so a `comment` element with both
 * `ref="A1"` and `ref="A1"` (impossible per spec) wouldn't match the
 * wrong block.
 */
function extractTextElementForComment(
  commentsXml: string,
  ref: string,
  authorId: string | undefined
): string | undefined {
  if (!ref) return undefined;
  const refEsc = escapeForRegex(ref);
  const idEsc = authorId !== undefined ? escapeForRegex(authorId) : undefined;
  const idAttrPart = idEsc !== undefined ? `\\s+authorId="${idEsc}"` : "";
  const re = new RegExp(`<comment\\s+ref="${refEsc}"${idAttrPart}[^>]*>([\\s\\S]*?)<\\/comment>`, "g");
  const m = re.exec(commentsXml);
  if (!m) return undefined;
  const inner = m[1] ?? "";
  const textRe = /<text\b[^>]*>([\s\S]*?)<\/text>|<text\s*\/>/;
  const tm = textRe.exec(inner);
  if (!tm) return undefined;
  return tm[0];
}

function escapeForRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
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
