/**
 * Public entry for the embeddable PPTX editor — see the docstring in
 * `./docx.ts` for why these wrappers re-export from `apps/web/app/...`
 * via the `@/` alias and how the `.css` side-effect import is
 * picked up by the `officeai-css-inject` esbuild plugin.
 *
 * The PPTX editor renders shape geometry via inline JSX styles, so
 * it doesn't depend on `prosemirror-view` (DOCX) or PDF.js's
 * text-layer CSS (PDF). It does need the editor design tokens
 * (`--divider`, `--surface`, …) and `.pptx-comment-flash` from
 * `styles.generated.css` so the comments-sidebar locate-flash and
 * surface-coloured chrome render.
 */
import "../styles.generated.css";

export { PptxEditor } from "@/pptx-editor/PptxEditor";
export type { EmbeddedEditorProps } from "../contract";
