/**
 * Public entry for the embeddable XLSX editor — see the docstring in
 * `./docx.ts` for why these wrappers re-export from `apps/web/app/...`
 * via the `@/` alias and how the `*.css` side-effect imports are
 * picked up by the `officeai-css-inject` esbuild plugin.
 *
 * `styles.generated.css` carries `.xlsx-grid-cell { position:
 * absolute; … }` (extracted from `apps/web/app/globals.css` by
 * `scripts/build-styles.mjs`). Without it the grid renders but
 * every cell collapses to `position: static` and the spreadsheet
 * stacks at one pixel.
 */
import "../styles.generated.css";

export { XlsxEditor } from "@/xlsx-editor/XlsxEditor";
export type { EmbeddedEditorProps } from "../contract";
