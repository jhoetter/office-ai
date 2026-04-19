/**
 * Client-side PPTX export helpers — render slides to SVG, optionally
 * rasterize to PNG via an offscreen canvas, and bundle them into a
 * zip. Slide-range parsing for the rich Export dialog also lives here.
 *
 * The pure data path (snapshot → SVG strings) reuses the
 * already-built renderer in `@officeai/pptx/renderer`. Rasterization
 * is browser-only because we draw via `<canvas>`; we feature-detect
 * `OffscreenCanvas` and fall back to a hidden DOM canvas otherwise.
 */

import JSZip from "jszip";
import { slideToSvgString, type SvgRenderCtx } from "@officeai/pptx/renderer";
import type { PptxSnapshot, Slide } from "@officeai/pptx";

/* ── slide range parsing ─────────────────────────────────────────── */

/**
 * Parse a slide-range mini-DSL like `1,3-5,7` into a sorted, deduped
 * list of 0-based slide indices. Empty / whitespace input means "all
 * slides". Out-of-range entries are dropped silently — the caller
 * owns user-facing error messages.
 */
export function parseSlideRange(
  input: string,
  totalSlides: number
): ReadonlyArray<number> {
  const trimmed = input.trim();
  if (trimmed.length === 0) {
    return Array.from({ length: totalSlides }, (_, i) => i);
  }
  const out = new Set<number>();
  for (const piece of trimmed.split(",")) {
    const part = piece.trim();
    if (part.length === 0) continue;
    const dash = part.indexOf("-");
    if (dash === -1) {
      const n = Number(part);
      if (Number.isInteger(n) && n >= 1 && n <= totalSlides) out.add(n - 1);
      continue;
    }
    const start = Number(part.slice(0, dash));
    const end = Number(part.slice(dash + 1));
    if (!Number.isInteger(start) || !Number.isInteger(end)) continue;
    const lo = Math.min(start, end);
    const hi = Math.max(start, end);
    for (let i = lo; i <= hi; i += 1) {
      if (i >= 1 && i <= totalSlides) out.add(i - 1);
    }
  }
  return Array.from(out).sort((a, b) => a - b);
}

/* ── SVG rendering ───────────────────────────────────────────────── */

export interface SlideRenderOptions {
  /** 1-based scale multiplier applied when rasterizing to PNG. The
   * raw slide is treated as 960px wide — `scale: 2` produces 1920px,
   * matching the `@2x`/`@3x` retina convention. */
  readonly scale?: 1 | 2 | 3;
  /** Slide indices (0-based) to include. Defaults to every slide. */
  readonly indices?: ReadonlyArray<number>;
}

interface RenderedSlide {
  readonly index: number;
  readonly svg: string;
  readonly viewBoxWidth: number;
  readonly viewBoxHeight: number;
}

function renderSlide(
  snapshot: PptxSnapshot,
  slide: Slide,
  index: number
): RenderedSlide {
  const ctx: SvgRenderCtx = {
    slideSize: snapshot.root.slideSize,
    theme: snapshot.root.themeDefault,
    renderPlaceholderHints: false,
  };
  const svg = slideToSvgString(slide, ctx);
  return {
    index,
    svg,
    viewBoxWidth: emuToPx(snapshot.root.slideSize.cxEmu),
    viewBoxHeight: emuToPx(snapshot.root.slideSize.cyEmu),
  };
}

interface IndexedSlide {
  readonly slide: Slide;
  readonly index: number;
}

function selectSlides(
  snapshot: PptxSnapshot,
  indices?: ReadonlyArray<number>
): ReadonlyArray<IndexedSlide> {
  const all = snapshot.root.slides;
  if (!indices || indices.length === 0) {
    return all.map((slide, index) => ({ slide, index }));
  }
  const allowed = new Set(indices);
  const out: IndexedSlide[] = [];
  all.forEach((slide, index) => {
    if (allowed.has(index)) out.push({ slide, index });
  });
  return out;
}

/**
 * Bundle one SVG per slide into a zip. The SVG is the same one the
 * editor canvas renders (with placeholder hints suppressed for a
 * clean export look).
 */
export async function snapshotToSvgZip(
  snapshot: PptxSnapshot,
  options: SlideRenderOptions = {}
): Promise<Blob> {
  const zip = new JSZip();
  const slides = selectSlides(snapshot, options.indices);
  const pad = padWidth(snapshot.root.slides.length);
  for (const { slide, index } of slides) {
    const rendered = renderSlide(snapshot, slide, index);
    const name = `slide-${String(index + 1).padStart(pad, "0")}.svg`;
    zip.file(name, rendered.svg);
  }
  return await zip.generateAsync({ type: "blob", mimeType: "application/zip" });
}

/**
 * Bundle one PNG per slide into a zip. Rasterized client-side via an
 * offscreen canvas; throws if the environment doesn't support
 * canvas (e.g. SSR — call from a click handler, not from render).
 */
export async function snapshotToPngZip(
  snapshot: PptxSnapshot,
  options: SlideRenderOptions = {}
): Promise<Blob> {
  if (typeof window === "undefined") {
    throw new Error("PNG export is only available in the browser.");
  }
  const scale = options.scale ?? 2;
  const slides = selectSlides(snapshot, options.indices);
  const zip = new JSZip();
  const pad = padWidth(snapshot.root.slides.length);
  for (const { slide, index } of slides) {
    const rendered = renderSlide(snapshot, slide, index);
    const png = await rasterizeSvgToPng(rendered, scale);
    const name = `slide-${String(index + 1).padStart(pad, "0")}.png`;
    zip.file(name, png);
  }
  return await zip.generateAsync({ type: "blob", mimeType: "application/zip" });
}

async function rasterizeSvgToPng(
  rendered: RenderedSlide,
  scale: number
): Promise<Blob> {
  const targetWidth = Math.max(1, Math.round(rendered.viewBoxWidth * scale));
  const targetHeight = Math.max(1, Math.round(rendered.viewBoxHeight * scale));
  // Wrap the SVG with explicit width/height so `<img>` rasterizes at
  // the right pixel dimensions regardless of viewBox setup.
  const sized = rendered.svg.replace(
    /<svg(\s)/i,
    `<svg width="${targetWidth}" height="${targetHeight}"$1`
  );
  const svgBlob = new Blob([sized], { type: "image/svg+xml;charset=utf-8" });
  const url = URL.createObjectURL(svgBlob);
  try {
    const img = await loadImage(url);
    const canvas = document.createElement("canvas");
    canvas.width = targetWidth;
    canvas.height = targetHeight;
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas 2D context unavailable.");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, targetWidth, targetHeight);
    ctx.drawImage(img, 0, 0, targetWidth, targetHeight);
    return await new Promise<Blob>((resolve, reject) => {
      canvas.toBlob(
        (blob) => {
          if (blob) resolve(blob);
          else reject(new Error("Failed to encode PNG."));
        },
        "image/png"
      );
    });
  } finally {
    URL.revokeObjectURL(url);
  }
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("Failed to load slide SVG."));
    img.src = url;
  });
}

/* ── helpers ─────────────────────────────────────────────────────── */

/** EMU → px at the SVG renderer's canonical 96 DPI. */
function emuToPx(emu: number): number {
  return Math.round((emu / 914400) * 96);
}

function padWidth(total: number): number {
  return String(Math.max(1, total)).length;
}
