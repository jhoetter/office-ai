/**
 * EMU ↔ pixel conversions. Pure functions, no DOM.
 *
 * EMU = English Metric Unit (1/914400 of an inch). PowerPoint stores all
 * geometry in EMU, so the renderer's coordinate system inside the SVG is
 * also EMU; the browser scales the SVG to its container via CSS.
 *
 * For HTML overlays (contenteditable text) we convert to pixels at the
 * conventional 96 DPI.
 */

export const EMU_PER_INCH = 914400;
export const DEFAULT_DPI = 96;
export const EMU_PER_PX_AT_96DPI = EMU_PER_INCH / DEFAULT_DPI; // 9525

export function emuToPx(emu: number, dpi: number = DEFAULT_DPI): number {
  return (emu * dpi) / EMU_PER_INCH;
}

export function pxToEmu(px: number, dpi: number = DEFAULT_DPI): number {
  return Math.round((px * EMU_PER_INCH) / dpi);
}

/** A:rPr `sz` is in hundredths of a point. Convert to CSS pixels at given DPI. */
export function fontSizeHundredthsToPx(sz: number, dpi: number = DEFAULT_DPI): number {
  // 1 pt = 1/72 inch.
  const pt = sz / 100;
  return (pt * dpi) / 72;
}
