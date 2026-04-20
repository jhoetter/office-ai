import type { PdfEngineGlyphRun, PdfEngineTextItem } from "../types.js";

/**
 * Lift PDF.js text items (or any backend's `PdfEngineTextItem[]`) into
 * per-character glyph runs with bounding boxes in PDF user-space.
 *
 * PDF.js's `getTextContent()` reports a single `width`/`height` per
 * item, not per character. To produce the per-glyph rects the
 * Structured Text service needs (line bucketing, column detection,
 * search-hit highlighting), we walk the item's transform and
 * subdivide its width across the characters.
 *
 * The subdivision is "even" — each glyph gets `width / chars.length`.
 * That's an approximation, but the structured text consumer only
 * needs glyph bbox precision down to "which line is this on" /
 * "what's the leftmost / rightmost glyph in the line" / "merge the
 * selected glyphs into per-line quads". For all of those, even
 * subdivision lands within ~2pt of the true rect, which is well
 * within the line-bucket tolerance and the on-screen highlight
 * tolerance at typical zoom. (PDFium can later replace this with the
 * exact `FPDFText_GetCharBox` rect; the contract is the same.)
 *
 * The `transform` matrix is `[a, b, c, d, e, f]` where `(e, f)` is
 * the text origin and the `[[a,b],[c,d]]` 2x2 maps the text space
 * basis to user space. For the dominant case of horizontal,
 * un-skewed text:
 *   a > 0, d > 0, b == c == 0
 *   font size  ≈ |d|
 *   advance    ≈ width / a  (in text-space) but `width` is already in
 *                user-space, so the per-char advance is simply
 *                `width / chars.length` along the run direction.
 *
 * For rotated text we project the per-char advance vector along the
 * basis: `dx = width * a / |a|`, `dy = width * b / |a|`. Vertical
 * writing (`a==0, b!=0, c!=0, d==0`) falls out naturally.
 */
export function buildGlyphRuns(
  items: ReadonlyArray<PdfEngineTextItem>,
): PdfEngineGlyphRun[] {
  const out: PdfEngineGlyphRun[] = [];
  for (const it of items) {
    if (!it.str) continue;
    const chars = [...it.str]; // surrogate-pair safe
    const n = chars.length;
    if (n === 0) continue;

    const [a, b, c, d, e, f] = it.transform;
    // Projection of the unit advance onto user-space.
    const advanceMag = Math.hypot(a, b) || 1;
    const ux = a / advanceMag;
    const uy = b / advanceMag;
    const totalWidth = it.width || advanceMag * n;
    const charAdvance = totalWidth / n;

    // Cap-height is approximately the font size, which equals the
    // magnitude of the basis y-vector `(c, d)`.
    const fontHeight = Math.hypot(c, d) || it.height || advanceMag;

    // Vertical step (perpendicular to the advance, +y in user-space)
    // gives us the bbox top from the baseline. PDF user-space has
    // origin bottom-left so "up" is +y.
    const perpX = -uy;
    const perpY = ux;

    const glyphs: Array<readonly [number, number, number, number]> = [];
    let minX = Number.POSITIVE_INFINITY;
    let minY = Number.POSITIVE_INFINITY;
    let maxX = Number.NEGATIVE_INFINITY;
    let maxY = Number.NEGATIVE_INFINITY;

    for (let i = 0; i < n; i++) {
      const ax0 = e + ux * (charAdvance * i);
      const ay0 = f + uy * (charAdvance * i);
      const ax1 = e + ux * (charAdvance * (i + 1));
      const ay1 = f + uy * (charAdvance * (i + 1));

      // Quad corners: baseline ↔ baseline + ascent vector.
      const tx = perpX * fontHeight;
      const ty = perpY * fontHeight;
      const xs = [ax0, ax1, ax0 + tx, ax1 + tx];
      const ys = [ay0, ay1, ay0 + ty, ay1 + ty];
      const x1 = Math.min(...xs);
      const x2 = Math.max(...xs);
      const y1 = Math.min(...ys);
      const y2 = Math.max(...ys);
      glyphs.push([x1, y1, x2, y2] as const);
      if (x1 < minX) minX = x1;
      if (y1 < minY) minY = y1;
      if (x2 > maxX) maxX = x2;
      if (y2 > maxY) maxY = y2;
    }

    out.push({
      chars: it.str,
      glyphs,
      baselineY: f,
      bbox: [minX, minY, maxX, maxY] as const,
      fontHeight,
      fontKey: it.fontName ?? "default",
      // PDF.js delivers items in visual order. Right-to-left scripts
      // surface as runs whose baseline advance vector points left
      // (negative `a` after a rotate). Use `a` sign as a proxy.
      dir: a < 0 ? "rtl" : "ltr",
      hasEol: Boolean(it.hasEol),
    });
  }
  return out;
}
