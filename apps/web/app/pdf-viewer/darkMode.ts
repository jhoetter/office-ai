/**
 * Dark-mode rendering helper for raster PDF pages.
 *
 * We use a CSS filter on the canvas (`filter: invert(1) hue-rotate(180deg)`)
 * because it is instant, per-frame free, and uniformly handles every
 * glyph regardless of the underlying engine. The trade-off is that
 * photographs invert too — acceptable for a binary toggle the user
 * controls. The previous "smart" per-pixel pass was removed for being
 * confusing and inconsistent across page kinds.
 */
export function darkModeCssFilter(enabled: boolean): string {
  return enabled ? "invert(1) hue-rotate(180deg)" : "none";
}
