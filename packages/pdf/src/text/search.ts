/**
 * Glyph-precise search across a `StructuredPage`.
 *
 * Builds a per-page concatenated string that exactly mirrors what
 * `serializeReadingOrder()` produces (so character offsets returned
 * here line up with the offsets in `page.text`). For each regex
 * match we walk back through the offset → glyph index map and
 * group the matched glyphs into per-line union rects in PDF
 * user-space.
 *
 * Match rects can be passed straight to the viewer's
 * `PdfHighlight` / `gotoMatch()` overlay.
 */
import type { PdfRect } from "../model/types.js";
import type { StructuredGlyph, StructuredLine, StructuredPage } from "./structured.js";

export interface StructuredSearchHit {
  readonly start: number;
  readonly end: number;
  readonly match: string;
  /** Per-line union rects of the matched glyphs (PDF user-space). */
  readonly rects: ReadonlyArray<PdfRect>;
}

interface CharSlot {
  readonly char: string;
  /** When non-null, this slot maps to a real glyph on a line. */
  readonly glyph?: { line: StructuredLine; glyph: StructuredGlyph };
}

export function findInStructuredPage(
  page: StructuredPage,
  re: RegExp,
  expectedPlain: string
): ReadonlyArray<StructuredSearchHit> | null {
  const slots = buildCharSlots(page);
  // The serializer-derived plain text we receive (`page.text`) is
  // built from the same structured blocks, so it should match
  // character-for-character. If a downstream mutation has changed
  // the text without touching `structured`, bail out so the agent
  // falls back to the legacy string-only search instead of producing
  // misaligned offsets.
  if (slotsToString(slots) !== expectedPlain) return null;

  const out: StructuredSearchHit[] = [];
  re.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(expectedPlain)) !== null) {
    const matchStart = m.index;
    const matchEnd = m.index + m[0].length;
    const rects = collectRects(slots, matchStart, matchEnd);
    out.push({ start: matchStart, end: matchEnd, match: m[0], rects });
    if (m[0].length === 0) re.lastIndex++;
  }
  return out;
}

function buildCharSlots(page: StructuredPage): CharSlot[] {
  const slots: CharSlot[] = [];
  let lastColumn = -1;
  let firstBlock = true;
  for (const block of page.blocks) {
    if (firstBlock) {
      firstBlock = false;
    } else if (block.columnIndex !== lastColumn && lastColumn !== -1) {
      // Mirrors `serializeReadingOrder`: a blank-line column break
      // followed by the standard "\n\n" between blocks.
      pushString(slots, "\n\n\n");
    } else {
      pushString(slots, "\n\n");
    }
    lastColumn = block.columnIndex;

    block.lines.forEach((line, lineIndex) => {
      if (lineIndex > 0) pushString(slots, "\n");
      appendLine(slots, line);
    });
  }
  return slots;
}

function appendLine(slots: CharSlot[], line: StructuredLine): void {
  // Re-derive the same text the structured builder produced (with
  // its inter-glyph " " padding) so the reading-order offsets stay
  // bit-identical with `serializeReadingOrder`.
  const expected = line.text;
  let g = 0;
  for (let i = 0; i < expected.length; i++) {
    const ch = expected[i];
    const next = line.glyphs[g];
    if (next && next.char === ch) {
      slots.push({ char: ch, glyph: { line, glyph: next } });
      g++;
    } else {
      slots.push({ char: ch });
    }
  }
}

function pushString(slots: CharSlot[], s: string): void {
  for (const ch of s) slots.push({ char: ch });
}

function slotsToString(slots: ReadonlyArray<CharSlot>): string {
  let s = "";
  for (const slot of slots) s += slot.char;
  return s;
}

function collectRects(slots: ReadonlyArray<CharSlot>, start: number, end: number): PdfRect[] {
  const byLine = new Map<StructuredLine, { x1: number; y1: number; x2: number; y2: number }>();
  for (let i = start; i < end; i++) {
    const slot = slots[i];
    if (!slot?.glyph) continue;
    const { line, glyph } = slot.glyph;
    const cur = byLine.get(line);
    const [gx1, gy1, gx2, gy2] = glyph.bbox;
    if (!cur) {
      byLine.set(line, { x1: gx1, y1: gy1, x2: gx2, y2: gy2 });
    } else {
      if (gx1 < cur.x1) cur.x1 = gx1;
      if (gy1 < cur.y1) cur.y1 = gy1;
      if (gx2 > cur.x2) cur.x2 = gx2;
      if (gy2 > cur.y2) cur.y2 = gy2;
    }
  }
  const out: PdfRect[] = [];
  // Emit rects in top-to-bottom (descending y) order so the viewer
  // can flash them sequentially without sorting.
  const sorted = [...byLine.values()].sort((a, b) => b.y2 - a.y2);
  for (const r of sorted) out.push([r.x1, r.y1, r.x2, r.y2] as PdfRect);
  return out;
}
