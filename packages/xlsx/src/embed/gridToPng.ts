import { encodePng } from "@officeai/core";
import type { EmbeddedGrid } from "./buildEmbeddedXlsx.js";

/**
 * Tiny pure-JS rasterizer that turns an `EmbeddedGrid` into a PNG
 * preview image suitable for OLE object preview shapes. Word and
 * PowerPoint will display this image while the embed is "cold" (not
 * activated by double-click); after activation Excel takes over and
 * paints the live cells.
 *
 * The renderer doesn't try to be pretty — it draws:
 *   - white background
 *   - light gray grid lines on cell boundaries
 *   - a slightly darker header band for row 0 (series names)
 *   - 3-pixel "letter blocks" for each character so the user can see
 *     there is actual content (no font rendering / Bezier glyphs
 *     because we have no font available in pure JS without pulling
 *     in `node-canvas` or `@resvg/resvg`)
 *
 * The output dimensions are bounded to keep the embedded preview a
 * reasonable size; rows/cols beyond the bounds are truncated visually
 * (the underlying xlsx is unchanged).
 */
export interface GridToPngOptions {
  readonly cellWidth?: number;
  readonly cellHeight?: number;
  readonly maxWidth?: number;
  readonly maxHeight?: number;
}

export interface GridToPngResult {
  readonly bytes: Uint8Array;
  readonly width: number;
  readonly height: number;
}

const DEFAULT_CELL_W = 80;
const DEFAULT_CELL_H = 22;
const DEFAULT_MAX_W = 640;
const DEFAULT_MAX_H = 480;

const COLOR_BG = [255, 255, 255, 255] as const;
const COLOR_GRID = [217, 217, 217, 255] as const;
const COLOR_HEADER_BG = [242, 242, 242, 255] as const;
const COLOR_TEXT = [64, 64, 64, 255] as const;

export function gridToPng(grid: EmbeddedGrid, opts: GridToPngOptions = {}): GridToPngResult {
  const cellW = Math.max(8, opts.cellWidth ?? DEFAULT_CELL_W);
  const cellH = Math.max(8, opts.cellHeight ?? DEFAULT_CELL_H);
  const maxW = Math.max(cellW, opts.maxWidth ?? DEFAULT_MAX_W);
  const maxH = Math.max(cellH, opts.maxHeight ?? DEFAULT_MAX_H);

  const rows = Math.max(1, grid.length);
  const cols = Math.max(
    1,
    grid.reduce((m, r) => Math.max(m, r.length), 0)
  );
  const visibleCols = Math.min(cols, Math.floor(maxW / cellW));
  const visibleRows = Math.min(rows, Math.floor(maxH / cellH));
  const width = Math.max(cellW, visibleCols * cellW);
  const height = Math.max(cellH, visibleRows * cellH);

  const pixels = new Uint8Array(width * height * 4);
  fillRect(pixels, width, height, 0, 0, width, height, COLOR_BG);

  for (let c = 0; c < visibleCols; c++) {
    fillRect(pixels, width, height, 0, 0, 1, height, COLOR_GRID);
    fillRect(pixels, width, height, c * cellW, 0, 1, height, COLOR_GRID);
  }
  fillRect(pixels, width, height, width - 1, 0, 1, height, COLOR_GRID);
  for (let r = 0; r < visibleRows; r++) {
    fillRect(pixels, width, height, 0, r * cellH, width, 1, COLOR_GRID);
  }
  fillRect(pixels, width, height, 0, height - 1, width, 1, COLOR_GRID);

  if (visibleRows > 0) {
    fillRect(pixels, width, height, 1, 1, width - 2, cellH - 2, COLOR_HEADER_BG);
  }

  for (let r = 0; r < visibleRows; r++) {
    const row = grid[r] ?? [];
    for (let c = 0; c < visibleCols; c++) {
      const value = row[c];
      if (value === undefined || value === null) continue;
      const text = String(value);
      drawTextStub(pixels, width, height, c * cellW + 4, r * cellH + 6, cellW - 8, cellH - 12, text);
    }
  }

  const bytes = encodePng({ width, height, rgba: pixels });
  return { bytes, width, height };
}

/** Fill a solid rectangle in the RGBA buffer. */
function fillRect(
  pixels: Uint8Array,
  width: number,
  height: number,
  x: number,
  y: number,
  w: number,
  h: number,
  color: readonly [number, number, number, number]
): void {
  const x1 = Math.max(0, Math.min(width, x));
  const y1 = Math.max(0, Math.min(height, y));
  const x2 = Math.max(0, Math.min(width, x + w));
  const y2 = Math.max(0, Math.min(height, y + h));
  for (let yy = y1; yy < y2; yy++) {
    for (let xx = x1; xx < x2; xx++) {
      const off = (yy * width + xx) * 4;
      pixels[off] = color[0];
      pixels[off + 1] = color[1];
      pixels[off + 2] = color[2];
      pixels[off + 3] = color[3];
    }
  }
}

/**
 * Render a row of "letter blocks" — one tiny 3px-wide block per
 * character — so the preview is visually distinguishable when the cell
 * actually has text in it. We can't do real glyph rendering without a
 * font binary; this is enough to convey "there's content here" while
 * staying dependency-free.
 */
function drawTextStub(
  pixels: Uint8Array,
  width: number,
  height: number,
  x: number,
  y: number,
  maxW: number,
  maxH: number,
  text: string
): void {
  const charW = 4;
  const charH = Math.min(maxH, 8);
  const maxChars = Math.max(0, Math.floor(maxW / charW));
  const visible = text.slice(0, maxChars);
  for (let i = 0; i < visible.length; i++) {
    const ch = visible.charCodeAt(i);
    if (ch === 32) continue;
    fillRect(pixels, width, height, x + i * charW, y, charW - 1, charH, COLOR_TEXT);
  }
}
