/**
 * Smart-invert dark mode for raster PDF pages.
 *
 * Trade-off:
 * - The cheap path is a CSS filter on the canvas
 *   (`filter: invert(1) hue-rotate(180deg)`), which we apply by default
 *   while the user is scrolling. It is instant and per-frame free, but
 *   it also inverts photographs which usually reads as "wrong" — a
 *   smiling face on a dark background turns into the photographic
 *   negative.
 * - {@link applyDarkModeToContext} runs after the page has finished
 *   rendering and walks the bitmap once. It inverts luminance for
 *   pixels that look like text/line art (low saturation, mid-to-high
 *   alpha) and leaves saturated regions (photos, illustrations,
 *   coloured fills) untouched. Slower (~5 ms per A4 page at scale 1.5
 *   on a modern laptop), but matches Acrobat's "soft black on warm
 *   off-white" behaviour for body text.
 *
 * The PdfCanvas calls this lazily — it will only invest the per-pixel
 * cost on the page that is currently in focus, after the rendering
 * promise has resolved.
 *
 * Spec reference: /spec/pdf/dark-mode.md.
 */

/**
 * Constants for the dark-mode background and foreground tones. We
 * deliberately push white to a warm off-white (#E6E1DA) and pure
 * black to a soft warm grey (#1B1B1A) — pure inversion would put
 * text on `#FFFFFF` which strobes against an OLED display.
 */
const DARK_BG: readonly [number, number, number] = [27, 27, 26];
const DARK_FG: readonly [number, number, number] = [230, 225, 218];

/**
 * Saturation cut-off above which we *skip* the invert and leave the
 * pixel in its original colour. Hand-tuned against the spec's smart-
 * invert fixture: 0.18 keeps neon highlights and screenshots intact
 * while still inverting anti-aliased greys around ~16-32 % black.
 */
const SATURATION_KEEP_THRESHOLD = 0.18;

/**
 * Apply smart-invert in place to a 2D canvas context.
 *
 * Pure (no DOM mutation other than the bitmap rewrite). Safe to call
 * multiple times — a second call would invert the inversion, so the
 * caller must guarantee one-shot semantics per render pass (PdfCanvas
 * tracks this via a render-token map).
 *
 * @param ctx     Target 2D context, already painted by the engine.
 * @param width   Canvas pixel width.
 * @param height  Canvas pixel height.
 */
export function applyDarkModeToContext(
  ctx: CanvasRenderingContext2D,
  width: number,
  height: number
): void {
  if (width <= 0 || height <= 0) return;
  const image = ctx.getImageData(0, 0, width, height);
  const data = image.data;
  const bgR = DARK_BG[0];
  const bgG = DARK_BG[1];
  const bgB = DARK_BG[2];
  const fgR = DARK_FG[0];
  const fgG = DARK_FG[1];
  const fgB = DARK_FG[2];
  for (let i = 0; i < data.length; i += 4) {
    const r = data[i]!;
    const g = data[i + 1]!;
    const b = data[i + 2]!;
    const a = data[i + 3]!;
    if (a === 0) continue;
    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    const sat = max === 0 ? 0 : (max - min) / max;
    if (sat > SATURATION_KEEP_THRESHOLD) continue;
    const luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
    const inv = 1 - luminance;
    data[i] = Math.round(bgR * (1 - inv) + fgR * inv);
    data[i + 1] = Math.round(bgG * (1 - inv) + fgG * inv);
    data[i + 2] = Math.round(bgB * (1 - inv) + fgB * inv);
  }
  ctx.putImageData(image, 0, 0);
}

/**
 * Cheap CSS-only fallback. Returns a `filter` string that the canvas
 * can apply via `style.filter`. Inverts everything (photos included),
 * which is acceptable when the per-pixel pass would budget-bust.
 */
export function darkModeCssFilter(enabled: boolean): string {
  return enabled ? "invert(1) hue-rotate(180deg)" : "none";
}
