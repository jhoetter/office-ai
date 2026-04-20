import type { PdfEngine } from "../types.js";

/**
 * PDFium-WASM backend (BSD-3-Clause via @embedpdf/pdfium).
 *
 * Status: structural placeholder. Spec'd in /spec/pdf/engine-strategy.md
 * §"PDFium-WASM fidelity fallback". The interface is the same as the
 * pdfjs backend — when wiring pdfium in, only this file changes.
 *
 * The choice to defer the real wiring is explicit: pdfium-wasm adds
 * ~10MB to the bundle and is only needed for fidelity-critical render
 * paths. We ship the engine-agnostic API + selector + tests today, and
 * land the real backend behind the same interface as a follow-up.
 */
export const pdfiumBackend: PdfEngine = {
  kind: "pdfium",
  async load(): Promise<never> {
    throw new Error(
      "pdf-engine: pdfium backend is registered but not yet wired. " +
        "See /spec/pdf/engine-strategy.md for the migration plan. " +
        "Use forceEngine: 'pdfjs' or omit forceEngine for the default backend."
    );
  },
};
