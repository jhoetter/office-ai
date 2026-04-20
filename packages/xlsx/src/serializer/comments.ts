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
    // P1: encode threaded-comment metadata (parentId, resolved,
    // createdAt) as `officeai-*` attributes. Excel preserves unknown
    // attributes when round-tripping the part, and our parser reads
    // them back via `parser/comments.ts`. Modern Excel clients
    // additionally surface these comments through their own
    // `xl/threadedComments/*` parts; we don't author those (yet).
    let attrs = `ref="${escapeXmlAttr(c.ref)}" authorId="${idx}"`;
    if (c.parentId) attrs += ` officeai-parentId="${escapeXmlAttr(c.parentId)}"`;
    if (c.resolved) attrs += ` officeai-resolved="1"`;
    if (c.createdAt) attrs += ` officeai-createdAt="${escapeXmlAttr(c.createdAt)}"`;
    parts.push(`<comment ${attrs}>`);
    // Prefer the captured opaque `<text>` element when the typed
    // `text` field is byte-identical to the flattened text inside
    // it. This preserves rich-text runs (multiple `<r>` children
    // with their own `<rPr>` font/color overrides), arbitrary
    // whitespace, and other run children we don't model. When the
    // user has actually edited the comment body the captured blob
    // is stale and we fall back to a single-run rebuild.
    if (c.textXml && flattenTextXml(c.textXml) === c.text) {
      parts.push(c.textXml);
    } else {
      parts.push("<text>");
      parts.push(`<r><t xml:space="preserve">${escapeXmlText(c.text)}</t></r>`);
      parts.push("</text>");
    }
    parts.push("</comment>");
  }
  parts.push("</commentList>");

  parts.push("</comments>");
  return parts.join("");
}

/**
 * Concatenate the text content from every `<t>` element inside a
 * captured `<text>` blob. Mirrors `extractCommentText` in
 * `parser/comments.ts` so the staleness check in
 * {@link serializeCommentsPart} agrees with what the parser
 * produced — without this, even an unedited comment would always
 * fall through to the single-run rebuild and lose its rich-text
 * formatting on every save.
 */
function flattenTextXml(textXml: string): string {
  let out = "";
  const re = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(textXml))) {
    out += unescapeXmlText(m[1] ?? "");
  }
  return out;
}

function unescapeXmlText(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function escapeXmlText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeXmlAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
