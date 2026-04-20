import type { PdfOutlineNode, PdfSnapshot } from "../model/types.js";

const renderOutline = (nodes: ReadonlyArray<PdfOutlineNode>, depth = 0): string => {
  if (nodes.length === 0) return "";
  const lines: string[] = [];
  for (const n of nodes) {
    const indent = "  ".repeat(depth);
    const pageRef = n.pageNumber !== undefined ? ` _(p. ${n.pageNumber})_` : "";
    lines.push(`${indent}- ${n.title}${pageRef}`);
    if (n.children.length > 0) lines.push(renderOutline(n.children, depth + 1));
  }
  return lines.join("\n");
};

/**
 * Markdown projection of a PDF — title metadata + outline + per-page
 * text. Designed for LLM consumption (mirrors snapshotToMarkdown for
 * docx / xlsx / pptx).
 */
export const snapshotToMarkdown = (snapshot: PdfSnapshot): string => {
  const out: string[] = [];
  const md = snapshot.root.metadata;
  if (md.title) out.push(`# ${md.title}`);
  if (md.author) out.push(`_by ${md.author}_`);
  if (md.subject) out.push(`> ${md.subject}`);

  if (snapshot.root.outline.length > 0) {
    out.push("\n## Outline\n");
    out.push(renderOutline(snapshot.root.outline));
  }

  out.push("\n## Pages\n");
  for (const page of snapshot.root.pages) {
    out.push(`\n### Page ${page.pageNumber}`);
    if (page.text.trim().length > 0) {
      out.push(page.text.trim());
    } else {
      out.push("_(no text layer — likely a scanned page; run OCR for searchable text)_");
    }
  }
  return out.join("\n");
};
