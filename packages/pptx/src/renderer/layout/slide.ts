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

/**
 * Padding fractions that surround the slide rectangle inside the editor's
 * "scratch canvas" (the gray artboard around the slide). 0.25 means a quarter
 * of the slide's width on each horizontal side, and a quarter of its height
 * on each vertical side. Exposed as constants so canvas pointer math, SVG
 * viewBox, and slide-card positioning all agree.
 */
export const STAGE_PAD_FRAC_X = 0.25;
export const STAGE_PAD_FRAC_Y = 0.25;

/**
 * Extended viewBox that includes a configurable scratch area around the
 * slide. Coordinate (0, 0) still sits at the top-left of the slide; negative
 * values map into the scratch margin. Used by the interactive editor canvas
 * so off-slide shapes remain visible while editing. Export and present mode
 * intentionally keep using {@link slideViewBox} so off-slide content is
 * cropped in delivered artifacts.
 */
export function slideStageViewBox(
  size: SlideSize,
  padFracX: number = STAGE_PAD_FRAC_X,
  padFracY: number = STAGE_PAD_FRAC_Y
): string {
  const cx = size.cxEmu * SVG_UNIT_PER_EMU;
  const cy = size.cyEmu * SVG_UNIT_PER_EMU;
  const padX = cx * padFracX;
  const padY = cy * padFracY;
  const x = round2(-padX);
  const y = round2(-padY);
  const w = round2(cx + padX * 2);
  const h = round2(cy + padY * 2);
  return `${x} ${y} ${w} ${h}`;
}

/**
 * Aspect ratio of the stage (slide + scratch padding). Useful for sizing the
 * outer pointer-capturing div so it preserves the same proportions as the
 * extended viewBox.
 */
export function stageAspectRatio(
  size: SlideSize,
  padFracX: number = STAGE_PAD_FRAC_X,
  padFracY: number = STAGE_PAD_FRAC_Y
): number {
  const w = size.cxEmu * (1 + padFracX * 2);
  const h = size.cyEmu * (1 + padFracY * 2);
  return w / h;
}

function round2(v: number): number {
  return Math.round(v * 100) / 100;
}
