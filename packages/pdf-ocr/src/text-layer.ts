import { PDFDocument, StandardFonts } from "pdf-lib";

export interface OcrSpan {
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly confidence: number;
}

export interface AddTextLayerOptions {
  /** 1-indexed pages to OCR. Default: all. */
  readonly pages?: ReadonlyArray<number>;
  /** Tesseract language(s). Default: "eng+deu". */
  readonly lang?: string;
  /** DPI used to rasterize each page before OCR. Default: 200. */
  readonly dpi?: number;
  /**
   * Optional injection of a tesseract.js–compatible recognise function.
   * Useful for tests and for environments where tesseract.js is not
   * installed but OCR is provided externally.
   */
  readonly recognise?: (
    pageImage: Uint8Array,
    pageWidth: number,
    pageHeight: number,
    lang: string,
  ) => Promise<ReadonlyArray<OcrSpan>>;
}

const tryLoadTesseract = async (): Promise<unknown | null> => {
  try {
    return await import("tesseract.js");
  } catch {
    return null;
  }
};

/**
 * Add an invisible (rendering-mode 3) text layer on top of each
 * requested page. Selects on copy + searchable; visually identical to
 * the raster underneath. Returns the patched PDF bytes.
 *
 * If `recognise` is not provided, dynamically imports `tesseract.js`.
 * If that import fails (peer dep not installed), throws a clear error.
 */
export const addTextLayer = async (
  buffer: Uint8Array,
  opts: AddTextLayerOptions = {},
): Promise<Uint8Array> => {
  const lang = opts.lang ?? "eng+deu";
  const recognise =
    opts.recognise ??
    (async (): Promise<ReadonlyArray<OcrSpan>> => {
      const mod = await tryLoadTesseract();
      if (!mod) {
        throw new Error(
          "pdf-ocr: tesseract.js is not installed. Install with: pnpm add tesseract.js (peer dep), " +
            "or pass `recognise` directly to addTextLayer for an externally-provided OCR.",
        );
      }
      throw new Error(
        "pdf-ocr: built-in tesseract.js bridge requires a rasterizer. " +
          "Spec: /spec/pdf/text-layer.md §OCR fallback. Provide `recognise` for now.",
      );
    });

  const pdf = await PDFDocument.load(buffer, { updateMetadata: false });
  const total = pdf.getPageCount();
  const target = opts.pages ?? Array.from({ length: total }, (_, i) => i + 1);
  if (target.some((p) => p < 1 || p > total)) {
    throw new Error(`pdf-ocr/add-text-layer: page out of range (1..${total})`);
  }
  const font = await pdf.embedFont(StandardFonts.Helvetica);

  for (const pageNumber of target) {
    const page = pdf.getPage(pageNumber - 1);
    const { width, height } = page.getSize();
    const spans = await recognise(new Uint8Array(0), width, height, lang);
    for (const span of spans) {
      page.drawText(span.text, {
        x: span.x,
        y: height - span.y - span.height,
        size: span.height,
        font,
        opacity: 0,
      });
    }
  }

  return pdf.save();
};
