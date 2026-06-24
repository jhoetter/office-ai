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
 * Worker / CMap bootstrap for embedded hosts:
 *
 * The Next.js host (apps/web) resolves `pdfjs-dist`'s worker with
 * `new URL("pdfjs-dist/build/pdf.worker.min.mjs", import.meta.url)`
 * and serves `/pdfjs/` from `apps/web/public/pdfjs/`. Embedding hosts
 * (e.g. embedding host, Vite-based) can't reliably do either — the bundled
 * chunk's `import.meta.url` doesn't yield a fetchable worker URL
 * inside `node_modules/.vite/deps/`, and the host has no static
 * `/pdfjs/` directory. We side-step both problems by setting the
 * documented globals at module-load time so PdfEditor's lazy
 * worker-init pass picks them up before the first parse.
 *
 * The CDN URLs are version-pinned to the `pdfjs-dist` dependency
 * declared in this package's `package.json`. Hosts that prefer to
 * self-host (offline / corporate-CDN / privacy reasons) can override
 * either global before importing this module — the `??=` keeps
 * pre-existing values intact.
 */
import "../styles.generated.css";

const PDFJS_VERSION = "4.10.38";
const CDN_BASE = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}`;

declare global {
  // eslint-disable-next-line no-var
  var __OFFICEAI_PDFJS_WORKER_SRC__: string | undefined;
  // eslint-disable-next-line no-var
  var __OFFICEAI_PDFJS_ASSETS_BASE__: string | undefined;
}

if (typeof globalThis !== "undefined") {
  globalThis.__OFFICEAI_PDFJS_WORKER_SRC__ ??= `${CDN_BASE}/build/pdf.worker.min.mjs`;
  globalThis.__OFFICEAI_PDFJS_ASSETS_BASE__ ??= `${CDN_BASE}/`;
}

export { PdfEditor } from "@/pdf-viewer/PdfEditor";
export type { EmbeddedEditorProps, PresenceUser } from "../contract";
