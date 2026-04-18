/**
 * EMU (English Metric Unit) conversions for OOXML drawings.
 *
 * OOXML stores all drawing dimensions in EMUs:
 *   - 914400 EMU = 1 inch
 *   - 96 DPI is the assumed CSS pixel grid
 *   - therefore 9525 EMU = 1 CSS pixel.
 *
 * We round to the nearest EMU on the way out and to the nearest tenth
 * of a pixel on the way back so a save-then-reload doesn't visibly
 * drift the image around the grid.
 */

export const EMU_PER_PX = 9525;

export function pxToEmu(px: number): number {
  return Math.round(px * EMU_PER_PX);
}

export function emuToPx(emu: number): number {
  return Math.round((emu / EMU_PER_PX) * 10) / 10;
}
