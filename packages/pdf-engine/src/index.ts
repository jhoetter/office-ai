/**
 * @officeai/pdf-engine — engine-agnostic PDF read/render API.
 *
 * Default backend: PDF.js (Apache 2.0). Optional fidelity fallback:
 * PDFium-WASM (BSD-3-Clause via @embedpdf/pdfium), lazy-loaded behind
 * the same interface (see /spec/pdf/engine-strategy.md).
 *
 * The interface here is engine-agnostic on purpose so the backend swap
 * is internal — see `selectEngine()` in src/select-engine.ts.
 */
export type {
  PdfEngine,
  PdfEngineKind,
  PdfEngineDocument,
  PdfEnginePage,
  PdfEnginePageInfo,
  PdfEngineRenderOptions,
  PdfEngineTextItem,
  PdfEngineTextContent,
  PdfEngineOutlineNode,
  PdfEngineAnnotationLite,
  PdfEngineFormFieldLite,
  PdfEngineMetadata,
  PdfEngineLoadOptions,
} from "./types.js";

export { loadDocument, getDefaultEngine, setDefaultEngine } from "./engine.js";
export { selectEngine, type EngineSelectionHints } from "./select-engine.js";
