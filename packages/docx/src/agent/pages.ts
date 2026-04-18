import type { BlockNode, DocxSnapshot, Paragraph, RunChild } from "../model/types.js";
import { chunkIntoPages, type PageChunk } from "../renderer/page-chunker.js";
import { paragraphPlainText } from "../commands/helpers.js";
import { snapshotToMarkdown } from "./markdown.js";

/**
 * P3.6 / W22-W24 — page-aware projections for the LLM/MCP surface.
 *
 * Wraps {@link chunkIntoPages} with metadata the agent layer needs:
 *
 * - **trigger** — what caused the page boundary (hard break, hint
 *   break, section break, or doc-start). Used by MCP clients to
 *   explain the layout to a user.
 * - **preview** — the first ≤120 chars of plain text on the page,
 *   whitespace-collapsed, so list views can show context without
 *   pulling the full markdown.
 *
 * These helpers are shared by the in-process agent (`DocxAgent.getPages`)
 * and the MCP server (`docx_get_pages` / `docx_get_page_text`) so both
 * stay in sync.
 */

export type PageTrigger =
  | "doc-start"
  | "page-break"
  | "last-rendered"
  | "section-break"
  | "measured-overflow";

export interface PageInfo {
  readonly pageNumber: number;
  /** Inclusive start index into `snapshot.root.body`. */
  readonly startBlockIndex: number;
  /** Exclusive end index into `snapshot.root.body`. */
  readonly endBlockIndex: number;
  /** Index into `body` of the SectionBreak whose geometry drives this page. */
  readonly sectionIndex: number;
  readonly trigger: PageTrigger;
  /** First ≤120 chars of body text on this page, whitespace-collapsed. */
  readonly preview: string;
}

const PREVIEW_MAX = 120;

/**
 * Produce one {@link PageInfo} per page chunk. Triggers are inferred
 * from the previous page's last block: a hard `<w:br w:type="page"/>`
 * yields `"page-break"`, a `<w:lastRenderedPageBreak/>` yields
 * `"last-rendered"`, and a `SectionBreak` block sitting at
 * `previousChunk.endBlockIndex - 1` (or `previousChunk.sectionIndex`
 * for empty sections) yields `"section-break"`. Page 1 is always
 * `"doc-start"`.
 */
export function getPageInfos(snapshot: DocxSnapshot): ReadonlyArray<PageInfo> {
  const chunks = chunkIntoPages(snapshot);
  const body = snapshot.root.body;
  return chunks.map((chunk, i) => {
    const trigger = inferTrigger(chunk, chunks[i - 1] ?? null, body);
    return {
      pageNumber: chunk.pageNumber,
      startBlockIndex: chunk.startBlock,
      endBlockIndex: chunk.endBlock,
      sectionIndex: chunk.sectionIndex,
      trigger,
      preview: buildPreview(chunk, body),
    };
  });
}

/**
 * Return the 1-based page number containing the body block at
 * `paragraphIndex`, or `null` when the index is out of range.
 *
 * Negative indices and indices ≥ `body.length` return `null`. Indices
 * inside an empty page (a section break with no preceding content)
 * resolve to that empty page.
 */
export function pageForParagraph(
  snapshot: DocxSnapshot,
  paragraphIndex: number
): number | null {
  if (paragraphIndex < 0 || paragraphIndex >= snapshot.root.body.length) return null;
  const chunks = chunkIntoPages(snapshot);
  for (let i = chunks.length - 1; i >= 0; i--) {
    const c = chunks[i];
    if (paragraphIndex >= c.startBlock && paragraphIndex < c.endBlock) return c.pageNumber;
    if (paragraphIndex >= c.startBlock && c.startBlock === c.endBlock) return c.pageNumber;
  }
  // Fallback: the paragraph index falls before page 1's startBlock —
  // can only happen for documents that begin with a SectionBreak as
  // the first body block. Round up to page 1 so the caller still
  // gets a sensible answer.
  return chunks.length > 0 ? chunks[0].pageNumber : null;
}

/**
 * Markdown projection of a single page. Returns `null` when
 * `pageNumber` is out of range (≤ 0 or > total pages).
 *
 * Internally builds a synthetic body slice and calls
 * {@link snapshotToMarkdown} on it so the formatting
 * (headings, lists, tables, comments) matches the full-document
 * output exactly.
 */
export function getPageMarkdown(snapshot: DocxSnapshot, pageNumber: number): string | null {
  const chunks = chunkIntoPages(snapshot);
  const chunk = chunks.find((c) => c.pageNumber === pageNumber);
  if (!chunk) return null;
  const slice = snapshot.root.body.slice(chunk.startBlock, chunk.endBlock);
  // Build a shallow snapshot scoped to this page. We keep the same
  // metadata (parts, comments, numbering, styles) so the markdown
  // projection still resolves heading levels and list formatting.
  const subSnapshot: DocxSnapshot = {
    ...snapshot,
    root: { ...snapshot.root, body: slice, comments: [] },
  };
  return snapshotToMarkdown(subSnapshot);
}

/**
 * Plain-text projection of a single page. Returns `null` when
 * `pageNumber` is out of range. Whitespace is preserved within
 * paragraphs and joined by blank lines between blocks — matches the
 * shape `docx_get_text` already returns for the full document with
 * `format: "text"`.
 */
export function getPagePlainText(snapshot: DocxSnapshot, pageNumber: number): string | null {
  const chunks = chunkIntoPages(snapshot);
  const chunk = chunks.find((c) => c.pageNumber === pageNumber);
  if (!chunk) return null;
  const lines: string[] = [];
  for (let i = chunk.startBlock; i < chunk.endBlock; i++) {
    const block = snapshot.root.body[i];
    if (block?.kind === "paragraph") lines.push(paragraphPlainText(block));
  }
  return lines.join("\n");
}

function inferTrigger(
  chunk: PageChunk,
  prev: PageChunk | null,
  body: ReadonlyArray<BlockNode>
): PageTrigger {
  if (!prev) return "doc-start";
  if (chunk.sectionIndex !== prev.sectionIndex) return "section-break";
  // The first block of the chunk is the one Word would render on the
  // new page. A hard break shows up *inside* that paragraph.
  const firstBlock = body[chunk.startBlock];
  if (firstBlock && firstBlock.kind === "paragraph") {
    if (paragraphHasChild(firstBlock, "page-break")) return "page-break";
    if (paragraphHasChild(firstBlock, "last-rendered-page-break")) return "last-rendered";
  }
  // Hard breaks can also live as the last leaf of the *previous*
  // page's last paragraph (the chunker keeps the carrying paragraph
  // on the new side, but if the break sits at end-of-paragraph the
  // chunker still treats it as hint-style for the next chunk).
  const lastPrev = body[prev.endBlock - 1];
  if (lastPrev && lastPrev.kind === "paragraph") {
    if (paragraphHasChild(lastPrev, "page-break")) return "page-break";
    if (paragraphHasChild(lastPrev, "last-rendered-page-break")) return "last-rendered";
  }
  return "measured-overflow";
}

function paragraphHasChild(p: Paragraph, kind: RunChild["kind"]): boolean {
  for (const inline of p.children) {
    if (inline.kind !== "run") continue;
    for (const child of inline.children) {
      if (child.kind === kind) return true;
    }
  }
  return false;
}

function buildPreview(chunk: PageChunk, body: ReadonlyArray<BlockNode>): string {
  let out = "";
  for (let i = chunk.startBlock; i < chunk.endBlock && out.length < PREVIEW_MAX; i++) {
    const block = body[i];
    if (block?.kind !== "paragraph") continue;
    const text = paragraphPlainText(block).replace(/\s+/g, " ").trim();
    if (text.length === 0) continue;
    if (out.length > 0) out += " · ";
    out += text;
  }
  out = out.trim();
  if (out.length > PREVIEW_MAX) out = out.slice(0, PREVIEW_MAX - 1) + "…";
  return out;
}
