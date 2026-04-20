/**
 * Structured Text — reading-order projection with glyph-level
 * bounding boxes for one PDF page.
 *
 * Lifts a flat array of `PdfEngineGlyphRun`s (which the engine
 * extracts in PDF stream order) into:
 *
 *   - Lines: bucketed by Y-baseline tolerance, sorted left→right
 *     within each line (right→left for RTL).
 *   - Columns: detected via the largest stable horizontal whitespace
 *     gap whose vertical extent covers most of the page; pages with
 *     no stable gap are treated as single-column.
 *   - Blocks: contiguous lines whose median font height + leading
 *     are within tolerance, classified as `paragraph` / `heading` /
 *     `list` via simple but battle-tested heuristics.
 *
 * Glyph indices are preserved through the whole pipeline so the
 * agent's `search()` and the viewer's `gotoMatch()` can map text
 * back to their bounding boxes without re-running PDF.js.
 *
 * Spec: /spec/pdf/text-layer.md.
 */
import type { PdfEngineGlyphRun } from "@officeai/pdf-engine";
import type { PdfRect } from "../model/types.js";

export interface StructuredGlyph {
  readonly char: string;
  /** Bounding box in PDF user-space (origin bottom-left). */
  readonly bbox: PdfRect;
}

export interface StructuredLine {
  readonly text: string;
  readonly glyphs: ReadonlyArray<StructuredGlyph>;
  /** Line bbox in PDF user-space. */
  readonly bbox: PdfRect;
  /** Median font height of glyphs in this line. */
  readonly fontHeight: number;
  /** Logical reading direction. */
  readonly dir: "ltr" | "rtl";
}

export type StructuredBlockKind = "paragraph" | "heading" | "list" | "caption";

export interface StructuredBlock {
  readonly kind: StructuredBlockKind;
  readonly lines: ReadonlyArray<StructuredLine>;
  readonly bbox: PdfRect;
  /** 0-based column index — 0 for single-column pages. */
  readonly columnIndex: number;
}

export interface StructuredPage {
  readonly pageWidth: number;
  readonly pageHeight: number;
  readonly blocks: ReadonlyArray<StructuredBlock>;
  /** Number of detected columns (1 for single-column pages). */
  readonly columnCount: number;
}

interface RawGlyph {
  readonly char: string;
  readonly bbox: PdfRect;
  readonly baselineY: number;
  readonly fontHeight: number;
  readonly dir: "ltr" | "rtl";
  readonly runIndex: number;
  readonly indexInRun: number;
}

const PRINTABLE = /\S/;

/**
 * Build a structured page from raw glyph runs. Pure function: no IO,
 * no engine dependency beyond the data contract.
 */
export function buildStructuredPage(
  runs: ReadonlyArray<PdfEngineGlyphRun>,
  pageWidth: number,
  pageHeight: number
): StructuredPage {
  const glyphs = flattenRuns(runs);
  if (glyphs.length === 0) {
    return { pageWidth, pageHeight, blocks: [], columnCount: 1 };
  }

  const columns = detectColumns(glyphs, pageWidth);
  const allBlocks: StructuredBlock[] = [];

  columns.forEach((columnGlyphs, columnIndex) => {
    const lines = bucketIntoLines(columnGlyphs);
    if (lines.length === 0) return;
    const blocks = groupLinesIntoBlocks(lines, columnIndex);
    for (const b of blocks) allBlocks.push(b);
  });

  // Column-major reading order: top-to-bottom within each column,
  // then move to the next column. PDF.js delivers stream order which
  // can be wildly different (e.g. interleaved between columns).
  allBlocks.sort((a, b) => {
    if (a.columnIndex !== b.columnIndex) return a.columnIndex - b.columnIndex;
    return b.bbox[3] - a.bbox[3]; // higher Y first (PDF origin bottom-left)
  });

  return {
    pageWidth,
    pageHeight,
    blocks: allBlocks,
    columnCount: columns.length,
  };
}

function flattenRuns(runs: ReadonlyArray<PdfEngineGlyphRun>): RawGlyph[] {
  const out: RawGlyph[] = [];
  runs.forEach((run, runIndex) => {
    const chars = [...run.chars];
    const n = Math.min(chars.length, run.glyphs.length);
    for (let i = 0; i < n; i++) {
      out.push({
        char: chars[i],
        bbox: run.glyphs[i] as PdfRect,
        baselineY: run.baselineY,
        fontHeight: run.fontHeight,
        dir: run.dir,
        runIndex,
        indexInRun: i,
      });
    }
  });
  return out;
}

/**
 * Column detection via projection-profile / largest-stable-gap.
 *
 * Build a horizontal occupancy histogram at 1pt resolution, then
 * scan for the widest vertical gap whose width is >= 18pt
 * (~conservative gutter) AND whose vertical extent covers >= 60%
 * of the inked page. If we find such a gap, split there; otherwise
 * the whole page is a single column.
 *
 * The 60% threshold rejects accidental gaps inside narrow figures
 * or short table rows; the 18pt minimum rejects normal inter-word
 * gaps even at large font sizes.
 *
 * For now we only split into at most two columns. Three+ column
 * layouts (newsprint) are rare in business documents and can be
 * added later by recursion.
 */
function detectColumns(glyphs: ReadonlyArray<RawGlyph>, pageWidth: number): RawGlyph[][] {
  if (glyphs.length < 30) return [glyphs.slice()];
  // Cheap occupancy mask: for each integer X column, true if any
  // glyph touches that column.
  const w = Math.max(1, Math.ceil(pageWidth));
  const occupied = new Uint8Array(w);
  let inkedYMin = Number.POSITIVE_INFINITY;
  let inkedYMax = Number.NEGATIVE_INFINITY;
  for (const g of glyphs) {
    const x1 = Math.max(0, Math.floor(g.bbox[0]));
    const x2 = Math.min(w - 1, Math.ceil(g.bbox[2]));
    for (let x = x1; x <= x2; x++) occupied[x] = 1;
    if (g.bbox[1] < inkedYMin) inkedYMin = g.bbox[1];
    if (g.bbox[3] > inkedYMax) inkedYMax = g.bbox[3];
  }
  const inkedHeight = Math.max(1, inkedYMax - inkedYMin);

  // Find the widest run of unoccupied columns, anchored well inside
  // the page (drop the outer 10% as margin).
  const margin = Math.floor(w * 0.1);
  let bestStart = -1;
  let bestLen = 0;
  let runStart = -1;
  for (let x = margin; x < w - margin; x++) {
    if (!occupied[x]) {
      if (runStart < 0) runStart = x;
    } else {
      if (runStart >= 0) {
        const len = x - runStart;
        if (len > bestLen) {
          bestLen = len;
          bestStart = runStart;
        }
        runStart = -1;
      }
    }
  }
  if (runStart >= 0) {
    const len = w - margin - runStart;
    if (len > bestLen) {
      bestLen = len;
      bestStart = runStart;
    }
  }

  if (bestLen < 18 || bestStart < 0) {
    return [glyphs.slice()];
  }

  // Verify the gap is "tall enough" — we look at the band centered
  // on the gap and check what fraction of inkedHeight is covered by
  // glyphs to its left vs right.
  const splitX = bestStart + bestLen / 2;
  let leftYMin = Number.POSITIVE_INFINITY;
  let leftYMax = Number.NEGATIVE_INFINITY;
  let rightYMin = Number.POSITIVE_INFINITY;
  let rightYMax = Number.NEGATIVE_INFINITY;
  for (const g of glyphs) {
    if (g.bbox[2] <= splitX) {
      if (g.bbox[1] < leftYMin) leftYMin = g.bbox[1];
      if (g.bbox[3] > leftYMax) leftYMax = g.bbox[3];
    } else if (g.bbox[0] >= splitX) {
      if (g.bbox[1] < rightYMin) rightYMin = g.bbox[1];
      if (g.bbox[3] > rightYMax) rightYMax = g.bbox[3];
    }
  }
  const leftHeight = leftYMax - leftYMin;
  const rightHeight = rightYMax - rightYMin;
  if (leftHeight < inkedHeight * 0.6 || rightHeight < inkedHeight * 0.6) {
    return [glyphs.slice()];
  }

  const left: RawGlyph[] = [];
  const right: RawGlyph[] = [];
  for (const g of glyphs) {
    const cx = (g.bbox[0] + g.bbox[2]) / 2;
    if (cx < splitX) left.push(g);
    else right.push(g);
  }
  return [left, right];
}

/**
 * Bucket glyphs into visual lines by baseline Y.
 *
 * Two glyphs share a line when:
 *   `|baselineY_a - baselineY_b| <= 0.5 * max(fontHeight_a, fontHeight_b)`
 *
 * Within a line, glyphs are sorted by X (LTR) or descending X (RTL,
 * detected by the line's dominant glyph direction). Lines emerge
 * sorted top-to-bottom (PDF origin is bottom-left, so descending Y).
 */
function bucketIntoLines(glyphs: ReadonlyArray<RawGlyph>): StructuredLine[] {
  if (glyphs.length === 0) return [];

  // Sort by descending baseline (top of page first).
  const sorted = [...glyphs].sort((a, b) => b.baselineY - a.baselineY);

  type Bucket = { baselineY: number; fontHeight: number; glyphs: RawGlyph[] };
  const buckets: Bucket[] = [];

  for (const g of sorted) {
    const tol = 0.5 * g.fontHeight;
    let placed = false;
    for (const b of buckets) {
      const diff = Math.abs(b.baselineY - g.baselineY);
      const t = Math.max(tol, 0.5 * b.fontHeight);
      if (diff <= t) {
        b.glyphs.push(g);
        // Update running average so progressive drift over many
        // glyphs doesn't fork into ghost lines on slightly slanted
        // page transforms.
        b.baselineY = (b.baselineY * (b.glyphs.length - 1) + g.baselineY) / b.glyphs.length;
        b.fontHeight = Math.max(b.fontHeight, g.fontHeight);
        placed = true;
        break;
      }
    }
    if (!placed) buckets.push({ baselineY: g.baselineY, fontHeight: g.fontHeight, glyphs: [g] });
  }

  return buckets.map((b) => buildLine(b.glyphs));
}

function buildLine(rawGlyphs: ReadonlyArray<RawGlyph>): StructuredLine {
  // Dominant direction of the line.
  let ltr = 0;
  let rtl = 0;
  for (const g of rawGlyphs) {
    if (g.dir === "rtl") rtl++;
    else ltr++;
  }
  const dir: "ltr" | "rtl" = rtl > ltr ? "rtl" : "ltr";
  const sorted = [...rawGlyphs].sort((a, b) => {
    const ax = (a.bbox[0] + a.bbox[2]) / 2;
    const bx = (b.bbox[0] + b.bbox[2]) / 2;
    return dir === "rtl" ? bx - ax : ax - bx;
  });

  const glyphs: StructuredGlyph[] = sorted.map((g) => ({ char: g.char, bbox: g.bbox }));
  let text = "";
  let prevRight: number | null = null;
  let prevFontHeight = sorted[0]?.fontHeight ?? 0;
  for (const g of sorted) {
    // Insert a single space when the inter-glyph gap is wider than
    // ~1/4 of an em. The engine's text items concatenate without a
    // separator so this is what gets us "Hello world" instead of
    // "Helloworld" when the source PDF emits one item per word.
    if (prevRight !== null) {
      const gap = dir === "rtl" ? prevRight - g.bbox[2] : g.bbox[0] - prevRight;
      if (gap > 0.25 * prevFontHeight && !/\s$/.test(text) && PRINTABLE.test(g.char)) {
        text += " ";
      }
    }
    text += g.char;
    prevRight = dir === "rtl" ? g.bbox[0] : g.bbox[2];
    prevFontHeight = g.fontHeight;
  }

  let x1 = Number.POSITIVE_INFINITY;
  let y1 = Number.POSITIVE_INFINITY;
  let x2 = Number.NEGATIVE_INFINITY;
  let y2 = Number.NEGATIVE_INFINITY;
  for (const g of sorted) {
    if (g.bbox[0] < x1) x1 = g.bbox[0];
    if (g.bbox[1] < y1) y1 = g.bbox[1];
    if (g.bbox[2] > x2) x2 = g.bbox[2];
    if (g.bbox[3] > y2) y2 = g.bbox[3];
  }
  const fontHeight = median(sorted.map((g) => g.fontHeight));

  return { text, glyphs, bbox: [x1, y1, x2, y2] as PdfRect, fontHeight, dir };
}

function median(xs: ReadonlyArray<number>): number {
  if (xs.length === 0) return 0;
  const s = [...xs].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 === 0 ? (s[mid - 1] + s[mid]) / 2 : s[mid];
}

/**
 * Group consecutive lines into blocks. A new block starts when:
 *   - The vertical gap to the previous line exceeds ~1.5× its
 *     leading (paragraph break).
 *   - The font height jumps by >20% (heading / body switch).
 *
 * Heading classification: any block whose median font height is
 * ≥1.2× the column's median.
 *
 * List classification: block whose first line starts with a bullet
 * glyph (`•`, `·`, `-`, `*`, `◦`, `▪`) or a numbered prefix
 * (`1.`, `1)`, `(1)`).
 */
const BULLET_GLYPHS = /^[\u2022\u2023\u25E6\u2043\u2219\u00B7\u25AA\-*]/;
const NUMBERED_PREFIX = /^(?:\(\d+\)|\d+[.)])/;

function groupLinesIntoBlocks(lines: ReadonlyArray<StructuredLine>, columnIndex: number): StructuredBlock[] {
  if (lines.length === 0) return [];
  const columnMedianFont = median(lines.map((l) => l.fontHeight));
  const blocks: StructuredBlock[] = [];
  let current: StructuredLine[] = [];
  let prev: StructuredLine | null = null;

  const flush = (): void => {
    if (current.length === 0) return;
    blocks.push(toBlock(current, columnMedianFont, columnIndex));
    current = [];
  };

  for (const line of lines) {
    if (prev) {
      const gap = prev.bbox[1] - line.bbox[3];
      const leading = Math.max(prev.fontHeight, line.fontHeight);
      const fontJump = Math.abs(line.fontHeight - prev.fontHeight) / Math.max(prev.fontHeight, 1);
      if (gap > 1.5 * leading || fontJump > 0.2) flush();
    }
    current.push(line);
    prev = line;
  }
  flush();
  return blocks;
}

function toBlock(
  lines: ReadonlyArray<StructuredLine>,
  columnMedianFont: number,
  columnIndex: number
): StructuredBlock {
  let x1 = Number.POSITIVE_INFINITY;
  let y1 = Number.POSITIVE_INFINITY;
  let x2 = Number.NEGATIVE_INFINITY;
  let y2 = Number.NEGATIVE_INFINITY;
  for (const l of lines) {
    if (l.bbox[0] < x1) x1 = l.bbox[0];
    if (l.bbox[1] < y1) y1 = l.bbox[1];
    if (l.bbox[2] > x2) x2 = l.bbox[2];
    if (l.bbox[3] > y2) y2 = l.bbox[3];
  }
  const blockFont = median(lines.map((l) => l.fontHeight));
  const firstText = lines[0].text.trimStart();
  let kind: StructuredBlockKind = "paragraph";
  if (blockFont >= columnMedianFont * 1.2) kind = "heading";
  else if (BULLET_GLYPHS.test(firstText) || NUMBERED_PREFIX.test(firstText)) kind = "list";
  return {
    kind,
    lines,
    bbox: [x1, y1, x2, y2] as PdfRect,
    columnIndex,
  };
}
