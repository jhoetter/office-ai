/**
 * Public entry for the embeddable PDF editor — see the docstring in
 * `./docx.ts` for why these wrappers re-export from `apps/web/app/...`
 * via the `@/` alias and how the `.css` side-effect import is
 * picked up by the `officeai-css-inject` esbuild plugin.
 *
 * `styles.generated.css` covers the shared editor tokens used by
 * the top-bar / toolbar chrome. PDF.js's text-layer stylesheet
 * (`apps/web/app/pdf-viewer/textLayer.css`) reaches the page
 * automatically — `PdfCanvas.tsx` already imports it, and the
 * `officeai-css-inject` plugin ensures every `.css` import in the
 * editor source graph self-injects on first load.
 *
 * NOTE: PdfEditor uses `pdfjs-dist` whose worker is loaded with
 * `import.meta.url`. Hosts must serve the `/pdfjs/` static asset
 * directory (see hof-os' integration guide) for rendering to work.
 */
import "../styles.generated.css";

export { PdfEditor } from "@/pdf-viewer/PdfEditor";
export type { EmbeddedEditorProps } from "../contract";
