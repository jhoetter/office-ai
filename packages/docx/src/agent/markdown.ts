import type { BlockNode, DocxComment, DocxSnapshot, Paragraph, Table } from "../model/types.js";
import { paragraphPlainText } from "../commands/helpers.js";
import { chunkIntoPages } from "../renderer/page-chunker.js";

export interface SnapshotToMarkdownOptions {
  /**
   * P3.6 / W22 — when true, segment the markdown output by page chunk
   * so LLMs can cite "page 3" without having to count themselves.
   * Each page is preceded by an HTML comment anchor
   * (`<!-- page N -->`) plus a `## Page N` heading; both are absent
   * when the option is false (default), preserving byte-identical
   * output for every existing CLI / MCP consumer.
   */
  readonly withPageSections?: boolean;
}

const HEADING_STYLES: Record<string, number> = {
  Title: 1,
  Heading1: 1,
  Heading2: 2,
  Heading3: 3,
  Heading4: 4,
  Heading5: 5,
  Heading6: 6,
};

/**
 * Project a DocxSnapshot to Markdown for AI consumption.
 *
 * - `SectionBreak` blocks become `---` separators.
 * - Paragraphs whose `styleId` matches `Title` / `Heading[1-6]` become
 *   `#` … `######` headings.
 * - Paragraphs with `numbering.numId` become numbered list items (`1.`),
 *   indented by `ilvl`. Bare `ListParagraph` style without numbering becomes
 *   a bullet (`-`).
 * - Tables are projected to GFM pipe tables (best effort: the table's first
 *   row becomes the header). When extraction fails the function falls back to
 *   the legacy `> [table preserved]` line and logs a warning.
 * - Comments, when present, are appended as a `## Comments` section listing
 *   the thread heads with the plain-text snippet of the parent paragraph.
 *
 * Pure: no I/O, no mutation. Safe to call from any context.
 */
export function snapshotToMarkdown(
  snapshot: DocxSnapshot,
  options?: SnapshotToMarkdownOptions
): string {
  const lines: string[] = [];
  const commentParents = buildCommentParentIndex(snapshot.root.body);

  // P3.6 / W22 — when page sections are requested, build a map from
  // body block index to the 1-based page number that block opens.
  // The map is sparse (one entry per chunk's `startBlock`) so the
  // emitter only injects a header at the start of each page.
  const pageStarts = options?.withPageSections ? buildPageStartIndex(snapshot) : null;

  for (let i = 0; i < snapshot.root.body.length; i++) {
    const block = snapshot.root.body[i];
    if (pageStarts) {
      const startPage = pageStarts.get(i);
      if (startPage !== undefined) {
        lines.push(`<!-- page ${startPage} -->`);
        lines.push(`## Page ${startPage}`);
        lines.push("");
      }
    }
    switch (block.kind) {
      case "paragraph":
        lines.push(paragraphToMarkdown(block));
        lines.push("");
        break;
      case "table": {
        // Prefer the typed model (introduced in P1.3 / W7). Fall back to the
        // legacy raw-subtree walker only if the typed shape is empty AND a
        // raw blob is still attached (back-compat for any code path that
        // somehow constructs a Table without rows but with bytes).
        const projected =
          block.rows.length > 0
            ? tableToMarkdownTyped(block)
            : block.raw
              ? tableToMarkdown(block.raw.subtree)
              : null;
        if (projected) {
          lines.push(...projected);
        } else {
          // Fall back to the legacy preservation hint. We log a single warning
          // per call site so noisy runs surface a hint without spamming.
          console.warn("snapshotToMarkdown: failed to extract table; falling back to placeholder");
          lines.push("> [table preserved]");
        }
        lines.push("");
        break;
      }
      case "section-break":
        lines.push("---");
        lines.push("");
        break;
      case "opaque-block":
        lines.push(`> [opaque block: ${block.raw.tag}]`);
        lines.push("");
        break;
      default: {
        const _exhaustive: never = block;
        void _exhaustive;
      }
    }
  }

  // Trailing empty page (e.g. a section break with no content after
  // it) — emit its header so the page count in the markdown matches
  // the chunker's view.
  if (pageStarts) {
    const trailingPage = pageStarts.get(snapshot.root.body.length);
    if (trailingPage !== undefined) {
      lines.push(`<!-- page ${trailingPage} -->`);
      lines.push(`## Page ${trailingPage}`);
      lines.push("");
    }
  }

  appendCommentsSection(lines, snapshot.root.comments, commentParents);

  return (
    lines
      .join("\n")
      .replace(/\n{3,}/g, "\n\n")
      .trimEnd() + "\n"
  );
}

/**
 * Map of body-block-index → 1-based page number for any block that
 * opens a new page. Empty pages (chunks whose `startBlock === endBlock`)
 * are recorded under that boundary index so the emitter still surfaces
 * a page header for them.
 */
function buildPageStartIndex(snapshot: DocxSnapshot): Map<number, number> {
  const out = new Map<number, number>();
  const chunks = chunkIntoPages(snapshot);
  for (const chunk of chunks) {
    out.set(chunk.startBlock, chunk.pageNumber);
  }
  return out;
}

function paragraphToMarkdown(p: Paragraph): string {
  const text = paragraphPlainText(p);
  const styleId = p.properties.styleId;
  if (styleId && HEADING_STYLES[styleId]) {
    const level = HEADING_STYLES[styleId];
    return `${"#".repeat(level)} ${text}`;
  }
  if (p.properties.numbering) {
    const indent = "  ".repeat(Math.max(0, p.properties.numbering.ilvl));
    return `${indent}1. ${text}`;
  }
  if (styleId === "ListParagraph") {
    return `- ${text}`;
  }
  return text;
}

/* ── Table extraction ────────────────────────────────────────────────────── */

/**
 * Extract a pipe-table from the opaque subtree of `<w:tbl>`. Returns `null`
 * when the table is empty or the structure doesn't look like a table — the
 * caller is expected to fall back to the placeholder line in that case.
 *
 * The walk is intentionally tolerant: anything that isn't `<w:tr>` is
 * skipped, anything that isn't `<w:tc>` inside a row is skipped, and cell
 * content is reduced to the concatenated `<w:t>` runs across every nested
 * paragraph (newlines collapsed to spaces so the row stays on one line).
 */
function tableToMarkdown(subtree: ReadonlyArray<unknown>): string[] | null {
  const rows: string[][] = [];
  for (const entry of elementsOf(subtree)) {
    const tag = tagOf(entry);
    if (tag !== "w:tr") continue;
    const cells: string[] = [];
    for (const cellEntry of elementsOf(childrenOf(entry, tag))) {
      if (tagOf(cellEntry) !== "w:tc") continue;
      cells.push(escapeCellText(extractText(childrenOf(cellEntry, "w:tc"))));
    }
    if (cells.length > 0) rows.push(cells);
  }
  if (rows.length === 0) return null;

  // Normalize column count: pad short rows so the resulting markdown is valid.
  const colCount = rows.reduce((m, r) => Math.max(m, r.length), 0);
  for (const r of rows) {
    while (r.length < colCount) r.push("");
  }

  const out: string[] = [];
  const header = rows[0];
  out.push(`| ${header.join(" | ")} |`);
  out.push(`| ${header.map(() => "---").join(" | ")} |`);
  for (let i = 1; i < rows.length; i++) {
    out.push(`| ${rows[i].join(" | ")} |`);
  }
  return out;
}

function escapeCellText(s: string): string {
  // Pipe characters break the table column boundary; collapse newlines.
  return s.replace(/\r?\n/g, " ").replace(/\|/g, "\\|").trim();
}

/**
 * Project a typed `Table` into a GFM pipe-table. Mirrors `tableToMarkdown`
 * (which still serves the legacy raw-subtree path) but reads cell text via
 * `paragraphPlainText` so it's resilient to mutations.
 */
function tableToMarkdownTyped(table: Table): string[] | null {
  const rows: string[][] = [];
  for (const row of table.rows) {
    const cells: string[] = [];
    for (const cell of row.cells) {
      const parts: string[] = [];
      for (const block of cell.body) {
        if (block.kind === "paragraph") parts.push(paragraphPlainText(block));
      }
      cells.push(escapeCellText(parts.join(" ")));
    }
    if (cells.length > 0) rows.push(cells);
  }
  if (rows.length === 0) return null;

  const colCount = rows.reduce((m, r) => Math.max(m, r.length), 0);
  for (const r of rows) {
    while (r.length < colCount) r.push("");
  }

  const out: string[] = [];
  const header = rows[0];
  out.push(`| ${header.join(" | ")} |`);
  out.push(`| ${header.map(() => "---").join(" | ")} |`);
  for (let i = 1; i < rows.length; i++) {
    out.push(`| ${rows[i].join(" | ")} |`);
  }
  return out;
}

/* ── Comment projection ──────────────────────────────────────────────────── */

/**
 * Map of `commentId` → plain-text snippet of the paragraph the comment is
 * anchored in. Built once per `snapshotToMarkdown` call; only top-level
 * threads appear in the output, but the index is keyed by every comment id
 * so replies can also be attributed to the correct anchor if a follow-up
 * format change wants them.
 */
function buildCommentParentIndex(body: ReadonlyArray<BlockNode>): Map<string, string> {
  const index = new Map<string, string>();
  for (const block of body) {
    if (block.kind !== "paragraph") continue;
    const text = paragraphPlainText(block);
    for (const child of block.children) {
      if (child.kind === "comment-range-start") {
        if (!index.has(child.commentId)) index.set(child.commentId, text);
      }
    }
  }
  return index;
}

function appendCommentsSection(
  lines: string[],
  comments: ReadonlyArray<DocxComment>,
  parents: ReadonlyMap<string, string>
): void {
  if (comments.length === 0) return;
  const heads = comments.filter((c) => c.parentId === undefined);
  if (heads.length === 0) return;

  lines.push("## Comments");
  lines.push("");
  for (const head of heads) {
    const body = commentBodyText(head);
    const parent = parents.get(head.id);
    const author = head.author || "unknown";
    const status = head.resolved ? " (resolved)" : "";
    const snippet = parent ? ` on "${truncate(parent, 60)}"` : "";
    lines.push(`- **${author}**${status}${snippet}: ${body}`);
    for (const reply of comments) {
      if (reply.parentId !== head.id) continue;
      lines.push(`  - **${reply.author || "unknown"}**: ${commentBodyText(reply)}`);
    }
  }
  lines.push("");
}

function commentBodyText(c: DocxComment): string {
  const parts: string[] = [];
  for (const b of c.body) {
    if (b.kind === "paragraph") parts.push(paragraphPlainText(b));
  }
  return parts.join(" ").trim();
}

function truncate(s: string, n: number): string {
  if (s.length <= n) return s;
  return s.slice(0, n - 1) + "…";
}

/* ── XML subtree helpers (preserveOrder) ─────────────────────────────────── */

const ATTR_KEY = ":@";

function elementsOf(siblings: ReadonlyArray<unknown>): Array<Record<string, unknown>> {
  const out: Array<Record<string, unknown>> = [];
  for (const s of siblings) {
    if (!s || typeof s !== "object" || Array.isArray(s)) continue;
    const obj = s as Record<string, unknown>;
    const keys = Object.keys(obj).filter((k) => k !== ATTR_KEY);
    if (keys.length === 1 && !keys[0].startsWith("?") && keys[0] !== "#text") {
      out.push(obj);
    }
  }
  return out;
}

function tagOf(entry: Record<string, unknown>): string {
  for (const k of Object.keys(entry)) {
    if (k !== ATTR_KEY) return k;
  }
  return "";
}

function childrenOf(entry: Record<string, unknown>, tag: string): ReadonlyArray<unknown> {
  const v = entry[tag];
  return Array.isArray(v) ? v : [];
}

/**
 * Recursively collect concatenated `<w:t>` text from a subtree. Used to
 * pull cell text without having to model every nested element typed.
 */
function extractText(siblings: ReadonlyArray<unknown>): string {
  let out = "";
  for (const s of siblings) {
    if (!s || typeof s !== "object" || Array.isArray(s)) continue;
    const obj = s as Record<string, unknown>;
    const tag = tagOf(obj);
    if (!tag || tag === ATTR_KEY) continue;
    const kids = obj[tag];
    if (tag === "w:t") {
      // The `<w:t>` element wraps a `#text` node in preserveOrder mode.
      if (Array.isArray(kids)) {
        for (const k of kids) {
          if (k && typeof k === "object" && !Array.isArray(k)) {
            const t = (k as Record<string, unknown>)["#text"];
            if (t !== undefined) out += String(t);
          }
        }
      }
    } else if (Array.isArray(kids)) {
      // Insert a space across paragraph boundaries inside a cell so words
      // don't smash together.
      const before = out;
      out += extractText(kids);
      if (tag === "w:p" && before !== out && !out.endsWith(" ")) out += " ";
    }
  }
  return out;
}
