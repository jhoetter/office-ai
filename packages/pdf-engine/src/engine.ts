import type { PdfEngine, PdfEngineDocument, PdfEngineKind, PdfEngineLoadOptions } from "./types.js";

/**
 * Lazy backend registry. Backends register on first use to keep the
 * default bundle small (PDFium-WASM is heavy).
 */
type BackendLoader = () => Promise<PdfEngine>;

const backends = new Map<PdfEngineKind, BackendLoader>();
let defaultEngine: PdfEngineKind = "pdfjs";

export const registerBackend = (kind: PdfEngineKind, loader: BackendLoader): void => {
  backends.set(kind, loader);
};

export const setDefaultEngine = (kind: PdfEngineKind): void => {
  defaultEngine = kind;
};

export const getDefaultEngine = (): PdfEngineKind => defaultEngine;

/**
 * Load a PDF document via the requested (or default) engine. Backends
 * are dynamic-imported on first use; if the requested backend is not
 * registered, throws a clear error pointing at the registration path.
 */
export const loadDocument = async (
  buffer: Uint8Array,
  opts: PdfEngineLoadOptions = {}
): Promise<PdfEngineDocument> => {
  const requested: PdfEngineKind = opts.forceEngine ?? defaultEngine;
  let loader = backends.get(requested);

  if (!loader) {
    if (requested === "pdfjs") {
      loader = async () => (await import("./backends/pdfjs.js")).pdfjsBackend;
      backends.set("pdfjs", loader);
    } else if (requested === "pdfium") {
      loader = async () => (await import("./backends/pdfium.js")).pdfiumBackend;
      backends.set("pdfium", loader);
    } else {
      throw new Error(
        `pdf-engine: no backend registered for "${requested as string}". ` +
          `Call registerBackend(${JSON.stringify(requested)}, …) first.`
      );
    }
  }

  const engine = await loader();
  return engine.load(buffer, opts);
};
