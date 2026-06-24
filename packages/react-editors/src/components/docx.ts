/**
 * Public entry for the embeddable DOCX editor.
 *
 * The implementation lives in `apps/web/app/editor/DocxEditor.tsx` so
 * Next's editor route and the bundled-for-embedding host build both ship the
 * exact same component (single source of truth, no copy/paste drift).
 * The esbuild step in this package resolves the `@/` alias against
 * `apps/web/app` and aliases the lone `next/link` import in
 * `EditorTopBar.tsx` to a plain-`<a>` shim — see `build.mjs` and
 * `src/shims/next-link.tsx` for the wiring.
 *
 * The two `*.css` side-effect imports below are picked up by the
 * `officeai-css-inject` esbuild plugin (see `build.mjs`) which
 * rewrites every `.css` import in the bundle's graph into a
 * self-injecting `<style>` tag pushed into the host page's `<head>`
 * on first load. These mirror what `apps/web/app/layout.tsx` brings
 * in for the Next.js host:
 *
 *   - `styles.generated.css`  — the editor-specific subset of
 *     `apps/web/app/globals.css` (design tokens, `.prose-pm` page
 *     sheet rendering, `.xlsx-grid-cell`, etc.) extracted by
 *     `scripts/build-styles.mjs`. Required by every editor.
 *
 *   - `prosemirror-view/style/prosemirror.css` — ProseMirror's own
 *     runtime stylesheet (gap-cursor, atom-selectednode outline,
 *     hideselection caret-color). The DOCX editor is the only
 *     bundled editor that uses ProseMirror, so we import it here
 *     instead of polluting the XLSX/PPTX/PDF bundles.
 */
import "../styles.generated.css";
import "prosemirror-view/style/prosemirror.css";

export { DocxEditor } from "@/editor/DocxEditor";
export type { EmbeddedEditorProps, PresenceUser } from "../contract";
