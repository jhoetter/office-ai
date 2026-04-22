/**
 * Lightweight text / Markdown serializers for the DOCX snapshot.
 *
 * These intentionally cover the everyday subset (paragraphs, headings,
 * lists, links, runs with emphasis, line breaks, tables) and not the
 * long tail (footnotes, structured-document tags, frames, complex
 * fields). Anything we don't model degrades to its plain-text form so
 * the export is never empty.
 */

import type {
  BlockNode,
  DocxSnapshot,
  Hyperlink,
  InlineNode,
  Paragraph,
  Run,
  RunChild,
  Table,
  TableCell,
} from "@officeai/docx";

interface SerializeContext {
  readonly document: DocxSnapshot["root"];
}

/**
 * Plain text view of the document. Headings and paragraphs become
 * lines; list items get a bullet/number prefix; tables collapse to
 * tab-separated rows; runs concatenate without formatting.
 */
export function docxToText(snapshot: DocxSnapshot): string {
  const ctx: SerializeContext = { document: snapshot.root };
  const lines: string[] = [];
  for (const block of snapshot.root.body) {
    appendBlockAsText(block, ctx, lines, 0);
  }
  return (
    lines
      .join("\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim() + "\n"
  );
}

/**
 * GitHub-flavoured Markdown. Headings 1-6 use `#` prefixes, runs apply
 * `**` / `_` / `~~`, bullet lists become `-`, ordered lists become
 * `1.`, hyperlinks become `[label](href)`, tables become pipe tables.
 */
export function docxToMarkdown(snapshot: DocxSnapshot): string {
  const ctx: SerializeContext = { document: snapshot.root };
  const out: string[] = [];
  for (const block of snapshot.root.body) {
    appendBlockAsMarkdown(block, ctx, out);
  }
  return (
    out
      .join("\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim() + "\n"
  );
}

/* ── plain text ───────────────────────────────────────────────────── */

function appendBlockAsText(block: BlockNode, ctx: SerializeContext, out: string[], indent: number): void {
  switch (block.kind) {
    case "paragraph": {
      const text = paragraphToPlain(block);
      const prefix = listPrefix(block);
      const padding = "  ".repeat(indent);
      if (text.length === 0 && !prefix) {
        out.push("");
        return;
      }
      out.push(`${padding}${prefix}${text}`);
      return;
    }
    case "table": {
      for (const row of block.rows) {
        const cells = row.cells.map((c) => cellToPlain(c, ctx));
        out.push(cells.join("\t"));
      }
      out.push("");
      return;
    }
    case "section-break":
      out.push("");
      return;
    case "opaque-block":
      if (block.children) {
        for (const child of block.children) appendBlockAsText(child, ctx, out, indent);
      }
      return;
    case "wrapper-marker":
      return;
    default: {
      const _exhaustive: never = block;
      void _exhaustive;
      return;
    }
  }
}

function cellToPlain(cell: TableCell, ctx: SerializeContext): string {
  const lines: string[] = [];
  for (const block of cell.body) appendBlockAsText(block, ctx, lines, 0);
  return lines.join(" ").replace(/\s+/g, " ").trim();
}

function paragraphToPlain(p: Paragraph): string {
  const out: string[] = [];
  for (const child of p.children) appendInlineAsText(child, out);
  return out.join("");
}

function appendInlineAsText(inline: InlineNode, out: string[]): void {
  switch (inline.kind) {
    case "run":
      out.push(runToPlain(inline));
      return;
    case "hyperlink":
      for (const r of inline.children) out.push(runToPlain(r));
      return;
    case "comment-range-start":
    case "comment-range-end":
    case "comment-reference":
      return;
    case "revision":
      if (inline.revisionType === "del") return;
      for (const child of inline.children) appendInlineAsText(child, out);
      return;
    case "opaque-inline":
      return;
    default: {
      const _exhaustive: never = inline;
      void _exhaustive;
      return;
    }
  }
}

function runToPlain(run: Run): string {
  const out: string[] = [];
  for (const child of run.children) {
    out.push(runChildToPlain(child));
  }
  return out.join("");
}

function runChildToPlain(child: RunChild): string {
  switch (child.kind) {
    case "text":
      return child.isDelText ? "" : child.text;
    case "tab":
      return "\t";
    case "break":
      return "\n";
    case "page-break":
      return "\n\n";
    case "last-rendered-page-break":
      return "";
    case "page-number-field":
      return child.cachedText ?? "";
    case "drawing":
      return "";
    case "opaque":
      return "";
    case "embedded-spreadsheet":
      return "";
    case "footnote-ref":
      return "";
    default: {
      const _exhaustive: never = child;
      void _exhaustive;
      return "";
    }
  }
}

/* ── Markdown ─────────────────────────────────────────────────────── */

function appendBlockAsMarkdown(block: BlockNode, ctx: SerializeContext, out: string[]): void {
  switch (block.kind) {
    case "paragraph": {
      const heading = headingLevel(block);
      if (heading > 0) {
        const text = paragraphInlinesToMarkdown(block);
        out.push(`${"#".repeat(heading)} ${text}`.trimEnd());
        out.push("");
        return;
      }
      const numbered = isOrderedList(block);
      const isList = block.properties.numbering !== undefined;
      if (isList) {
        const indent = "  ".repeat(block.properties.numbering?.ilvl ?? 0);
        const marker = numbered ? "1." : "-";
        const text = paragraphInlinesToMarkdown(block);
        out.push(`${indent}${marker} ${text}`.trimEnd());
        return;
      }
      const text = paragraphInlinesToMarkdown(block);
      out.push(text);
      out.push("");
      return;
    }
    case "table":
      appendTableAsMarkdown(block, ctx, out);
      out.push("");
      return;
    case "section-break":
      out.push("");
      out.push("---");
      out.push("");
      return;
    case "opaque-block":
      if (block.children) {
        for (const child of block.children) appendBlockAsMarkdown(child, ctx, out);
      }
      return;
    case "wrapper-marker":
      return;
    default: {
      const _exhaustive: never = block;
      void _exhaustive;
      return;
    }
  }
}

function appendTableAsMarkdown(table: Table, ctx: SerializeContext, out: string[]): void {
  if (table.rows.length === 0) return;
  const rows = table.rows.map((row) => row.cells.map((cell) => cellToMarkdownLine(cell, ctx)));
  const headerCells = rows[0]!;
  const colCount = headerCells.length;
  const align = new Array<string>(colCount).fill("---");
  out.push(`| ${headerCells.join(" | ")} |`);
  out.push(`| ${align.join(" | ")} |`);
  for (let i = 1; i < rows.length; i += 1) {
    const cells = rows[i]!;
    out.push(`| ${cells.join(" | ")} |`);
  }
}

function cellToMarkdownLine(cell: TableCell, ctx: SerializeContext): string {
  const lines: string[] = [];
  for (const block of cell.body) appendBlockAsMarkdown(block, ctx, lines);
  return lines
    .filter((l) => l.length > 0 && l !== "---")
    .join(" ")
    .replace(/\s+/g, " ")
    .replace(/\|/g, "\\|")
    .trim();
}

function paragraphInlinesToMarkdown(p: Paragraph): string {
  const out: string[] = [];
  for (const child of p.children) appendInlineAsMarkdown(child, out);
  return out.join("");
}

function appendInlineAsMarkdown(inline: InlineNode, out: string[]): void {
  switch (inline.kind) {
    case "run":
      out.push(runToMarkdown(inline));
      return;
    case "hyperlink":
      out.push(hyperlinkToMarkdown(inline));
      return;
    case "comment-range-start":
    case "comment-range-end":
    case "comment-reference":
      return;
    case "revision":
      if (inline.revisionType === "del") return;
      for (const child of inline.children) appendInlineAsMarkdown(child, out);
      return;
    case "opaque-inline":
      return;
    default: {
      const _exhaustive: never = inline;
      void _exhaustive;
      return;
    }
  }
}

function runToMarkdown(run: Run): string {
  const text = runToPlain(run)
    // Escape characters that have meaning inside Markdown emphasis
    // runs. We keep newlines as Markdown soft breaks.
    .replace(/\\/g, "\\\\")
    .replace(/\*/g, "\\*")
    .replace(/_/g, "\\_")
    .replace(/`/g, "\\`")
    .replace(/\[/g, "\\[")
    .replace(/\]/g, "\\]");
  if (text.length === 0) return "";
  let wrapped = text;
  if (run.properties.bold) wrapped = `**${wrapped}**`;
  if (run.properties.italic) wrapped = `_${wrapped}_`;
  if (run.properties.strike) wrapped = `~~${wrapped}~~`;
  return wrapped;
}

function hyperlinkToMarkdown(link: Hyperlink): string {
  const label = link.children.map((r) => runToMarkdown(r)).join("");
  if (link.anchor) return `[${label}](#${link.anchor})`;
  // We don't resolve the relationship target here — keep the label.
  // A future iteration can resolve via the snapshot's relationship
  // map; for now a label-only link is more useful than a raw rId.
  return label;
}

/* ── helpers ──────────────────────────────────────────────────────── */

function headingLevel(p: Paragraph): number {
  const id = p.properties.styleId;
  if (!id) return 0;
  const match = /^Heading(\d)$/.exec(id);
  if (!match) return 0;
  const level = Number(match[1]);
  if (!Number.isFinite(level) || level < 1) return 0;
  return Math.min(6, level);
}

function isOrderedList(p: Paragraph): boolean {
  // We don't have access to the numbering definitions reliably from
  // the snapshot here — fall back to the conservative default (bullet)
  // when ambiguous. The numId is captured so a future enhancement can
  // look it up in `snapshot.document.numbering`.
  void p;
  return false;
}

function listPrefix(p: Paragraph): string {
  if (p.properties.numbering === undefined) return "";
  const indent = "  ".repeat(Math.max(0, p.properties.numbering.ilvl));
  return `${indent}- `;
}
