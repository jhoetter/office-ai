/**
 * Engine-agnostic types for the PDF read/render surface. Both PDF.js
 * and PDFium-WASM backends conform to this contract.
 */

export type PdfEngineKind = "pdfjs" | "pdfium";

export interface PdfEngineLoadOptions {
  /** Optional password for encrypted PDFs. */
  password?: string;
  /** Override engine selection (otherwise auto-picked via selectEngine). */
  forceEngine?: PdfEngineKind;
  /**
   * Base URL (or `file://` path for Node) under which PDF.js worker
   * assets can be fetched. The backend appends `/cmaps/` and
   * `/standard_fonts/` to this base. Strongly recommended in the
   * browser — without it, CJK/Arabic CMaps and Type-1/Type-3 fallback
   * fonts are silently dropped, which makes selectable text disappear
   * for whole categories of real-world PDFs.
   *
   * Trailing slash is normalised; pass `/pdfjs` or `/pdfjs/` either
   * way. Default: undefined (no asset loading; PDF.js falls back to
   * the embedded heuristics, with the well-known correctness
   * limitations).
   */
  assetsBase?: string;
}

export interface PdfEngineDocument {
  readonly engine: PdfEngineKind;
  readonly numPages: number;
  /** 1-indexed page accessor matching the PDF spec convention. */
  getPage(pageNumber: number): Promise<PdfEnginePage>;
  getMetadata(): Promise<PdfEngineMetadata>;
  getOutline(): Promise<PdfEngineOutlineNode[] | null>;
  getAttachments(): Promise<ReadonlyArray<{ name: string; data: Uint8Array }>>;
  /** Total estimated memory used by the document, in bytes. Best-effort. */
  estimatedBytes(): number;
  /** Releases the underlying engine handles. */
  destroy(): Promise<void>;
}

export interface PdfEnginePageInfo {
  readonly pageNumber: number;
  /** PDF user-units (typically 1/72 inch). */
  readonly width: number;
  readonly height: number;
  /** 0 / 90 / 180 / 270, page-level rotation. */
  readonly rotation: 0 | 90 | 180 | 270;
  /** Optional page label (e.g. "iv", "A-12") if defined in the PDF. */
  readonly label?: string;
}

export interface PdfEngineRenderOptions {
  /** CSS pixels per PDF user-unit. Default ≈ 1.0 (~72 DPI). */
  scale?: number;
  /** Target canvas to draw into. Required in browser; auto-created in Node. */
  canvas?: HTMLCanvasElement | OffscreenCanvas;
  /** Output format for headless renderers. Defaults to "png". */
  format?: "png" | "jpeg" | "webp";
  /** Target DPI shortcut. Overrides `scale` if both set. */
  dpi?: number;
  /**
   * Extra clockwise rotation to apply on top of the page's intrinsic
   * rotation, in 90° increments. Default 0. The engine bakes the
   * rotation into the rasterised bitmap and the canvas dimensions
   * (width/height swap at 90°/270°) — callers should size their
   * canvas using the rotated viewport, not the un-rotated page.
   */
  rotation?: 0 | 90 | 180 | 270;
}

export interface PdfEngineTextItem {
  readonly str: string;
  /** Transform: [a, b, c, d, e, f] in PDF user-space. */
  readonly transform: readonly [number, number, number, number, number, number];
  readonly width: number;
  readonly height: number;
  readonly fontName?: string;
  readonly hasEol?: boolean;
}

export interface PdfEngineTextContent {
  readonly items: ReadonlyArray<PdfEngineTextItem>;
  /** Concatenated plain text in reading order (best-effort). */
  readonly plain: string;
}

/**
 * One contiguous text run (typically corresponds to a single
 * `TextItem` from PDF.js or one block from PDFium's
 * `FPDFText_GetText`). The `glyphs` array provides per-character
 * bounding boxes in PDF user-space (origin bottom-left), which is
 * what the Structured Text service needs for reading-order grouping
 * and glyph-precise search highlights.
 */
export interface PdfEngineGlyphRun {
  readonly chars: string;
  /** Per-character rect in PDF user-space (origin bottom-left). */
  readonly glyphs: ReadonlyArray<readonly [number, number, number, number]>;
  /** Baseline Y in PDF user-space. */
  readonly baselineY: number;
  /** Run-level bounding box (axis-aligned union of glyph rects). */
  readonly bbox: readonly [number, number, number, number];
  /** Approximate cap-height (font ascent ~ size). */
  readonly fontHeight: number;
  /** Stable font key (e.g. `"Helvetica"`); falls back to "default". */
  readonly fontKey: string;
  /** Logical reading direction inferred from the matrix. */
  readonly dir: "ltr" | "rtl";
  /** True when the engine reports an end-of-line break after this run. */
  readonly hasEol: boolean;
}

/**
 * Native viewport descriptor passed to PDF.js's official `TextLayer`
 * class (and PDFium's equivalent renderer). Treated as opaque by
 * everything except the engine and the viewer's text-layer wrapper.
 */
export interface PdfEngineViewport {
  readonly scale: number;
  readonly rotation: 0 | 90 | 180 | 270;
  /** Canvas width in CSS pixels. */
  readonly width: number;
  /** Canvas height in CSS pixels. */
  readonly height: number;
  /** Engine-specific raw viewport (PDF.js `PageViewport`). */
  readonly raw: unknown;
}

export interface PdfEnginePage {
  readonly info: PdfEnginePageInfo;
  /** Render the page; returns rendered bytes for headless / canvas for browser. */
  render(opts?: PdfEngineRenderOptions): Promise<Uint8Array | undefined>;
  getTextContent(): Promise<PdfEngineTextContent>;
  /**
   * Per-character glyph runs with bounding boxes in PDF user-space.
   * Backbone of the structured text service (reading-order blocks,
   * column detection, glyph-precise search highlights).
   */
  getGlyphRuns(): Promise<ReadonlyArray<PdfEngineGlyphRun>>;
  /**
   * Native viewport for use with the engine's canonical text layer.
   * For PDF.js this wraps `PDFPageProxy.getViewport({ scale, rotation })`
   * and exposes the raw `PageViewport` so the official `TextLayer`
   * class can be constructed from the viewer.
   */
  getViewport(opts: { scale: number; rotation?: 0 | 90 | 180 | 270 }): PdfEngineViewport;
  /**
   * Source for PDF.js's `TextLayer` constructor. Returns either a
   * `ReadableStream` (preferred — streams glyphs as the worker emits
   * them) or a fully resolved `TextContent` payload. Treat as opaque.
   */
  getTextContentSource(opts?: { includeMarkedContent?: boolean }): Promise<unknown>;
  getAnnotations(): Promise<ReadonlyArray<PdfEngineAnnotationLite>>;
  getFormFields(): Promise<ReadonlyArray<PdfEngineFormFieldLite>>;
  destroy(): void;
}

export interface PdfEngineOutlineNode {
  readonly title: string;
  /** 1-indexed page destination, if resolvable. */
  readonly pageNumber?: number;
  /** External URI destination, mutually exclusive with pageNumber. */
  readonly uri?: string;
  readonly children: ReadonlyArray<PdfEngineOutlineNode>;
}

/**
 * Lightweight annotation projection sufficient for read-side tooling.
 * The full annotation editing model lives in @officeai/pdf-annotations.
 */
export interface PdfEngineAnnotationLite {
  readonly id: string;
  readonly subtype: string;
  /** PDF user-space rect: [x1, y1, x2, y2]. */
  readonly rect: readonly [number, number, number, number];
  readonly contents?: string;
  readonly author?: string;
  readonly url?: string;
  readonly destPage?: number;
}

export interface PdfEngineFormFieldLite {
  readonly id: string;
  readonly name: string;
  readonly type: "text" | "checkbox" | "radio" | "choice" | "button" | "signature" | "unknown";
  readonly value?: string | boolean;
  readonly options?: ReadonlyArray<string>;
  readonly readOnly: boolean;
  readonly required: boolean;
  readonly maxLength?: number;
  readonly multiline?: boolean;
  readonly password?: boolean;
  readonly pageNumber: number;
  readonly rect: readonly [number, number, number, number];
}

export interface PdfEngineMetadata {
  readonly title?: string;
  readonly author?: string;
  readonly subject?: string;
  readonly keywords?: string;
  readonly creator?: string;
  readonly producer?: string;
  readonly creationDate?: string;
  readonly modificationDate?: string;
  /** Best-effort PDF version, e.g. "1.7". */
  readonly pdfVersion?: string;
  /** PDF security: present if the document was/is encrypted. */
  readonly encryption?: { readonly hasUserPassword: boolean; readonly hasOwnerPassword: boolean };
  /** Whether the document is linearized for fast web view. */
  readonly linearized?: boolean;
}

/** The contract every backend must implement. */
export interface PdfEngine {
  readonly kind: PdfEngineKind;
  load(buffer: Uint8Array, opts?: PdfEngineLoadOptions): Promise<PdfEngineDocument>;
}
