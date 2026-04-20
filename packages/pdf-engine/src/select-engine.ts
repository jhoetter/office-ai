import type { PdfEngineKind } from "./types.js";

/**
 * Hints surfaced from a partial PDF parse that influence engine choice.
 * See /spec/pdf/engine-strategy.md §"Auto-fallback heuristics".
 */
export interface EngineSelectionHints {
  /** PDF declares one of: DeviceN / NChannel / Separation / Lab in /ColorSpace. */
  readonly hasUncommonColorSpace?: boolean;
  /** Document uses non-standard CMaps that PDF.js may glyph-substitute. */
  readonly hasCustomCMap?: boolean;
  /** Document uses Type3 fonts (PDF.js renders these inconsistently). */
  readonly hasType3Fonts?: boolean;
  /** Document linearization flag. */
  readonly linearized?: boolean;
  /** User-curated allowlist hit. */
  readonly inPdfiumAllowlist?: boolean;
  /** User explicitly chose high-fidelity rendering. */
  readonly userPrefersFidelity?: boolean;
}

/**
 * Pure function that picks a backend given hints. PDF.js is the default;
 * PDFium-WASM is opted into when fidelity-critical signals fire.
 *
 * Spec: /spec/pdf/engine-strategy.md
 */
export const selectEngine = (hints: EngineSelectionHints = {}): PdfEngineKind => {
  if (hints.userPrefersFidelity === true) return "pdfium";
  if (hints.inPdfiumAllowlist === true) return "pdfium";
  if (hints.hasUncommonColorSpace === true) return "pdfium";
  if (hints.hasType3Fonts === true) return "pdfium";
  if (hints.hasCustomCMap === true) return "pdfium";
  return "pdfjs";
};
