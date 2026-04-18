import type { SlideSize } from "../../model/types.js";
import { EMU_PER_PX_AT_96DPI } from "./units.js";

/**
 * SVG renderer coordinate system: 1 user unit ≈ 1 CSS pixel at 96 DPI.
 *
 * PowerPoint geometry is stored in EMU (1/914400″). A direct EMU viewBox
 * works mathematically but Chrome's text rendering quietly degrades for
 * `font-size` values in the high 5- and 6-digit range (the glyphs render
 * almost invisibly small even when the viewport scaling should up-size
 * them proportionally). Using a viewBox in pixel-equivalents keeps every
 * font-size in the well-tested ~10–80px range while preserving the
 * source EMU positions via a single outer `<g transform="scale(…)">`.
 */
export const SVG_UNIT_PER_EMU = 1 / EMU_PER_PX_AT_96DPI;

export function slideViewBox(size: SlideSize): string {
  const cx = round2(size.cxEmu * SVG_UNIT_PER_EMU);
  const cy = round2(size.cyEmu * SVG_UNIT_PER_EMU);
  return `0 0 ${cx} ${cy}`;
}

export function slideAspectRatio(size: SlideSize): number {
  return size.cxEmu / size.cyEmu;
}

function round2(v: number): number {
  return Math.round(v * 100) / 100;
}
