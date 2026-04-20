import type { PdfRect } from "../model/types.js";

/**
 * Reading-order serialisers for `StructuredPage`.
 *
 * `serializeReadingOrder` is the canonical replacement for the old
 * `text.plain` (which was just stream order). Backs:
 *   - `PdfPage.text` (the plain text consumers like search and AI
 *     prompts read).
 *   - Copy-as-text in the viewer.
 *
 * `serializeMarkdown` is the structured walker used by
 * `agent.toMarkdown()` and the right-click "Copy as Markdown".
 * Maps detected block kinds onto Markdown primitives.
 */
import type {
  StructuredBlock,
  StructuredLine,
  StructuredPage,
} from "./structured.js";

export function serializeReadingOrder(page: StructuredPage): string {
  const out: string[] = [];
  let lastColumn = -1;
  for (const block of page.blocks) {
    if (block.columnIndex !== lastColumn && lastColumn !== -1) {
      // Hard break between columns so AI prompts and grep hits
      // don't run column boundaries together.
      out.push("");
    }
    lastColumn = block.columnIndex;
    out.push(blockText(block));
  }
  return out.join("\n\n").trim();
}

export function serializeMarkdown(page: StructuredPage): string {
  const out: string[] = [];
  for (const block of page.blocks) {
    out.push(renderBlockAsMarkdown(block));
  }
  return out.join("\n\n").trim();
}

function blockText(block: StructuredBlock): string {
  return block.lines.map(joinLine).join("\n");
}

/**
 * Collect text from glyphs that fall inside any of the supplied
 * PDF-user-space regions, emitted in the page's reading order.
 *
 * Used by the viewer's copy-event interception: the user drags
 * across multiple columns / paragraphs, the browser would
 * otherwise concat the spans in stream order. We instead replay
 * the structured-page walk and only include glyphs whose bbox
 * intersects the selection.
 *
 * Region intersection is "any overlap" — close enough for selection
 * highlights produced by the browser's `Range.getClientRects()`.
 */
export function collectTextWithinRegions(
  page: StructuredPage,
  regions: ReadonlyArray<PdfRect>,
): string {
  if (regions.length === 0) return "";
  const out: string[] = [];
  let lastColumn = -1;
  for (const block of page.blocks) {
    const blockText = collectTextFromBlock(block, regions);
    if (blockText.length === 0) continue;
    if (block.columnIndex !== lastColumn && lastColumn !== -1) out.push("");
    lastColumn = block.columnIndex;
    out.push(blockText);
  }
  return out.join("\n\n").trim();
}

function collectTextFromBlock(
  block: StructuredBlock,
  regions: ReadonlyArray<PdfRect>,
): string {
  const lineStrings: string[] = [];
  for (const line of block.lines) {
    let lineText = "";
    let prevIncluded = false;
    for (const glyph of line.glyphs) {
      if (intersectsAny(glyph.bbox, regions)) {
        if (!prevIncluded && lineText.length > 0 && !/\s$/.test(lineText)) {
          lineText += " ";
        }
        lineText += glyph.char;
        prevIncluded = true;
      } else {
        prevIncluded = false;
      }
    }
    if (lineText.trim().length > 0) lineStrings.push(lineText);
  }
  if (lineStrings.length === 0) return "";
  // Soft-wrap inside a paragraph: join with single spaces so
  // pasted text doesn't carry visual line breaks. List blocks keep
  // hard breaks because each line is a separate item.
  if (block.kind === "list") return lineStrings.join("\n");
  return lineStrings.join(" ");
}

function intersectsAny(
  bbox: readonly [number, number, number, number],
  regions: ReadonlyArray<PdfRect>,
): boolean {
  for (const r of regions) {
    if (bbox[2] < r[0] || bbox[0] > r[2]) continue;
    if (bbox[3] < r[1] || bbox[1] > r[3]) continue;
    return true;
  }
  return false;
}

function renderBlockAsMarkdown(block: StructuredBlock): string {
  switch (block.kind) {
    case "heading": {
      // Approximate heading levels by font size relative to other
      // headings on the same page. Without document-wide context we
      // can't tell h1 from h2 reliably, so we emit `##` for any
      // detected heading — pleasant default that round-trips
      // through Markdown processors without surprising users.
      const text = block.lines.map(joinLine).join(" ");
      return `## ${text}`;
    }
    case "list": {
      return block.lines
        .map((l) => `- ${stripBullet(joinLine(l))}`)
        .join("\n");
    }
    case "caption":
    case "paragraph":
      return block.lines.map(joinLine).join(" ");
    default: {
      const _exhaustive: never = block.kind;
      return _exhaustive;
    }
  }
}

function joinLine(line: StructuredLine): string {
  return line.text;
}

const BULLET_PREFIX_RE = /^\s*[\u2022\u2023\u25E6\u2043\u2219\u00B7\u25AA\-*]\s*/;
const NUMBERED_PREFIX_RE = /^\s*(?:\(\d+\)|\d+[.)])\s*/;

function stripBullet(text: string): string {
  return text.replace(BULLET_PREFIX_RE, "").replace(NUMBERED_PREFIX_RE, "");
}
