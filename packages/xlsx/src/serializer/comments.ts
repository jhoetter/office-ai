import type { Comment } from "../model/types.js";

const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/**
 * Serialize a typed `(authors, comments)` pair to `xl/comments{N}.xml`.
 *
 * Round-trip is **semantic** with `parser/comments.ts`: rich-text runs
 * collapse to a single `<r><t>` per comment because the model only
 * holds plain text in P0. `xml:space="preserve"` is set on the run
 * text so leading/trailing whitespace survives the round-trip.
 *
 * Author indices are derived from the `authors` array position; the
 * handler is responsible for keeping `commentAuthors` aligned with
 * the indices referenced in `comments`.
 */
export function serializeCommentsPart(
  authors: ReadonlyArray<string>,
  comments: ReadonlyArray<Comment>
): string {
  const parts: string[] = [];
  parts.push(XML_DECL);
  parts.push(`<comments xmlns="${NS}">`);

  parts.push("<authors>");
  for (const a of authors) {
    parts.push(`<author>${escapeXmlText(a)}</author>`);
  }
  parts.push("</authors>");

  parts.push("<commentList>");
  const authorIdx = new Map<string, number>();
  for (let i = 0; i < authors.length; i++) {
    if (!authorIdx.has(authors[i])) authorIdx.set(authors[i], i);
  }
  for (const c of comments) {
    const idx = authorIdx.get(c.author) ?? 0;
    parts.push(`<comment ref="${escapeXmlAttr(c.ref)}" authorId="${idx}">`);
    parts.push("<text>");
    parts.push(`<r><t xml:space="preserve">${escapeXmlText(c.text)}</t></r>`);
    parts.push("</text>");
    parts.push("</comment>");
  }
  parts.push("</commentList>");

  parts.push("</comments>");
  return parts.join("");
}

function escapeXmlText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeXmlAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
