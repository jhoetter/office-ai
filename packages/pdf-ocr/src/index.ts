/**
 * @officeai/pdf-ocr — text-layer generation for scanned PDFs.
 *
 * Spec: /spec/pdf/text-layer.md §"OCR fallback".
 *
 * Strategy: per-page raster (via @officeai/pdf-engine) → tesseract.js
 * recognise() → invisible-text overlay drawn on the same page via
 * pdf-lib (rendering mode 3 = invisible) at the recognised baseline.
 *
 * `tesseract.js` is an OPTIONAL peer dep so the bundle stays small for
 * users that don't need OCR. Calling `addTextLayer` without it throws
 * a clear error pointing at the install command.
 */
export { addTextLayer, type AddTextLayerOptions, type OcrSpan } from "./text-layer.js";
