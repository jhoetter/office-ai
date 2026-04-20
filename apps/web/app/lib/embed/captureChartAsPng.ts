/**
 * Rasterise a `ChartOverlay` DOM node — either an `<svg>` rendered
 * via React, or a pre-rendered `<canvas>` — into a PNG so the chart
 * can ride the system clipboard as `image/png` for external apps and
 * inside the cross-format `xlsx-chart-image` envelope for in-app
 * pastes into DOCX (image) and PPTX (picture).
 *
 * The XLSX charts in this app render via `<svg>`. To produce a PNG
 * we serialise the SVG, blob-URL it, paint it into an offscreen
 * `<canvas>` of the chart's intrinsic CSS pixel size, and read the
 * `image/png` data URL back out. The base64 payload is returned
 * stripped of the `data:` prefix so callers can either decode it to
 * `Uint8Array` or re-wrap it for transport without manual splitting.
 *
 * Failure modes (rejected promise):
 *   - The browser fails to load the SVG-as-image (malformed XML or
 *     CORS-tainted external assets — the chart overlay never uses
 *     external images, but a guard remains for safety).
 *   - The canvas refuses to export (taint-flagged).
 */

export interface CaptureChartArgs {
  /**
   * Source node. Pass either the `<svg>` the chart rendered into,
   * or a pre-rasterised `<canvas>`. The wrapping `<div data-testid>`
   * isn't useful directly; callers should resolve to the inner SVG /
   * canvas first.
   */
  readonly source: SVGSVGElement | HTMLCanvasElement;
  /**
   * Output canvas size in CSS pixels. Defaults to the source node's
   * client size — explicit dimensions are useful for tests and for
   * callers that want to upscale (e.g. for retina exports).
   */
  readonly width?: number;
  readonly height?: number;
}

export interface CaptureChartResult {
  /** Base64 PNG, no `data:image/png;base64,` prefix. */
  readonly png: string;
  readonly width: number;
  readonly height: number;
}

export async function captureChartAsPng(args: CaptureChartArgs): Promise<CaptureChartResult> {
  const { source } = args;
  if (typeof window === "undefined") {
    throw new Error("captureChartAsPng requires a browser environment.");
  }

  if (isCanvas(source)) {
    const width = args.width ?? source.width;
    const height = args.height ?? source.height;
    return { png: stripDataUrlPrefix(source.toDataURL("image/png")), width, height };
  }

  const svg = source;
  const width = Math.max(1, Math.floor(args.width ?? (svg.clientWidth || svgAttrPx(svg, "width"))));
  const height = Math.max(1, Math.floor(args.height ?? (svg.clientHeight || svgAttrPx(svg, "height"))));

  const xml = new XMLSerializer().serializeToString(svg);
  // Guarantee the namespace is declared on the root — XMLSerializer
  // sometimes drops it when the serialised node was lifted out of a
  // React-managed tree, which causes Image() to refuse to load it.
  const withNs = xml.includes("xmlns=")
    ? xml
    : xml.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"');

  const blob = new Blob([withNs], { type: "image/svg+xml;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  try {
    const img = await loadImage(url);
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Could not acquire 2D context for chart capture.");
    ctx.drawImage(img, 0, 0, width, height);
    return { png: stripDataUrlPrefix(canvas.toDataURL("image/png")), width, height };
  } finally {
    URL.revokeObjectURL(url);
  }
}

function isCanvas(node: unknown): node is HTMLCanvasElement {
  if (!node || typeof node !== "object") return false;
  if (typeof globalThis.HTMLCanvasElement === "function" && node instanceof globalThis.HTMLCanvasElement) {
    return true;
  }
  return typeof (node as { toDataURL?: unknown }).toDataURL === "function";
}

function svgAttrPx(svg: SVGSVGElement, attr: "width" | "height"): number {
  const raw = svg.getAttribute(attr);
  if (!raw) return 0;
  const n = Number.parseFloat(raw);
  return Number.isFinite(n) ? n : 0;
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("Failed to decode SVG into an image for chart capture."));
    img.src = url;
  });
}

function stripDataUrlPrefix(dataUrl: string): string {
  const comma = dataUrl.indexOf(",");
  return comma >= 0 ? dataUrl.slice(comma + 1) : dataUrl;
}
