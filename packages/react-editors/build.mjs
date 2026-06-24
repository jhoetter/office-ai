#!/usr/bin/env node
/**
 * Build the publishable artifacts for `@officeai/react-editors`.
 *
 * Architecture (intentionally pragmatic):
 *
 *  - Phase-1 entry points (`blanks/*`, `mime`, `contract`, `index`) are
 *    plain TypeScript living under `src/` and compiled by `tsc` (via
 *    `pnpm --filter ... run build:types` further down). They re-export
 *    workspace packages (`@officeai/{docx,xlsx,pptx,pdf}`) and have no
 *    Next coupling — straightforward.
 *
 *  - Phase-1.5 entry points (`components/{docx,xlsx,pptx,pdf}`) are
 *    *bundled by esbuild* from the existing
 *    `apps/web/app/{editor,xlsx-editor,pptx-editor,pdf-viewer}/*Editor.tsx`
 *    sources via the `@/` path alias. Reasons:
 *
 *      1. The editor source is ~11k LOC across the four formats plus
 *         a large shared `apps/web/app/lib/{shell,realtime,...}`
 *         subtree. Physically relocating it into the package would
 *         duplicate it (apps/web still needs it for its own routes)
 *         and burn days resolving imports.
 *
 *      2. The only Next-specific leak in the bundled graph is
 *         `next/link` in `lib/shell/EditorTopBar.tsx`. We alias that
 *         import to `src/shims/next-link.tsx` (a plain `<a>`), so the
 *         bundle has zero Next runtime dependency.
 *
 *      3. We externalize React, the workspace packages, and the
 *         heavy third-party deps (pdfjs-dist, prosemirror-*, lucide,
 *         jszip, yjs, y-websocket). They reach the consumer through
 *         pnpm's normal dependency resolution + the shameful-hoist
 *         step that the auto-release workflow runs (so the tarball
 *         that a downstream host's postinstall fetches has every external
 *         resolvable from the deploy root, even with Vite's
 *         `resolve.preserveSymlinks: true`).
 *
 * The output layout matches the `exports` block in `package.json` so
 * subpath imports (`@officeai/react-editors/components/xlsx`) resolve
 * cleanly in both Vite and a Node 22+ ESM consumer.
 */

import esbuild from "esbuild";
import { copyFileSync, mkdirSync, readdirSync, readFileSync, rmSync } from "node:fs";
import { createHash } from "node:crypto";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";
import path from "node:path";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const REPO_ROOT = path.resolve(__dirname, "..", "..");
const APPS_WEB_APP = path.join(REPO_ROOT, "apps", "web", "app");
const PKG_ROOT = __dirname;
const SRC = path.join(PKG_ROOT, "src");
const DIST = path.join(PKG_ROOT, "dist");
const NEXT_LINK_SHIM = path.join(SRC, "shims", "next-link.tsx");

// Externals = packages the bundle expects the HOST'S bundler (Vite,
// Next, webpack…) to resolve at runtime instead of inlining their
// source. Two distinct buckets:
//
//   1. The seven "format" workspace packages (@officeai/{core,docx,
//      xlsx,pptx,pdf,pdf-engine,pdf-annotations}) — large, ship
//      pre-built `dist/*.js`, declared as direct deps so pnpm + the
//      shameful-hoist step puts them at the deploy root. Inlining
//      them would balloon the bundle and break dedup with the
//      headless `@officeai/react-editors/blanks` entries.
//
//   2. React + heavy third-party libs (pdfjs-dist, prosemirror-*,
//      lucide, jszip, yjs, y-websocket) — same reason: they're
//      already in the host's dep graph (most hosts pull React from
//      somewhere), the bundler will dedupe, and inlining drags
//      multi-MB worker code into every chunk.
//
// NOT externalized:
//   - The five "shell" workspace packages (@officeai/{ui,
//     text-formatting,comments,realtime,design-tokens}) — they ship
//     RAW `.ts` from `./src/` (no `dist/`). Vite refuses to
//     transpile `.ts` files inside `node_modules`, so externalizing
//     them yields "Failed to resolve import './types'" the moment
//     the host loads an editor. Bundling them inline (~50 KB per
//     editor) sidesteps the whole problem and keeps the host's Vite
//     config clean.
const EXTERNAL = [
  "react",
  "react-dom",
  "react-dom/client",
  "react/jsx-runtime",
  "react/jsx-dev-runtime",
  "@officeai/docx",
  "@officeai/xlsx",
  "@officeai/pptx",
  "@officeai/pdf",
  "@officeai/core",
  "@officeai/pdf-engine",
  "@officeai/pdf-annotations",
  "pdfjs-dist",
  "pdfjs-dist/legacy/build/pdf.mjs",
  "lucide-react",
  "jszip",
  "yjs",
  "y-websocket",
  "prosemirror-model",
  "prosemirror-state",
  "prosemirror-view",
  "prosemirror-keymap",
  "prosemirror-commands",
  "prosemirror-history",
  "prosemirror-schema-basic",
  "prosemirror-schema-list",
  "prosemirror-tables",
  "prosemirror-inputrules",
  "prosemirror-transform",
];

rmSync(DIST, { recursive: true, force: true });
mkdirSync(DIST, { recursive: true });

// Step 1 — derive `dist/styles.css` and `src/styles.generated.css`
// from `apps/web/app/globals.css`. The `src/` mirror is what each
// component entry `import`s so esbuild can bundle the CSS into a
// sibling `dist/components/<format>.css` (loader: ".css" -> "css"
// below). Failing fast here keeps the CSS-missing class of bugs from
// shipping a broken tarball.
const stylesScript = path.join(PKG_ROOT, "scripts", "build-styles.mjs");
const stylesProc = spawnSync(process.execPath, [stylesScript], { stdio: "inherit" });
if (stylesProc.status !== 0) {
  console.error(`[react-editors:build] build-styles.mjs failed (exit ${stylesProc.status})`);
  process.exit(stylesProc.status ?? 1);
}

/**
 * esbuild plugin: rewrite every `.css` import in the bundle's graph
 * into a tiny self-injecting JS module that appends a `<style>` tag
 * to the host page's `<head>` on first load.
 *
 * Why a plugin instead of `loader: { ".css": "text" }` + a hand-rolled
 * one-stylesheet-only injector:
 *
 *   The Next.js host (`apps/web`) gets every editor stylesheet for
 *   free via Next's automatic CSS-import discovery — `layout.tsx`
 *   imports `prosemirror-view/style/prosemirror.css` and
 *   `./globals.css`, `pdf-viewer/PdfCanvas.tsx` imports
 *   `./textLayer.css`, and Next bundles + injects them all. Our
 *   esbuild bundle did NOT — the previous wiring only injected
 *   `globals.css` (via `build-styles.mjs` + a manual injector), and
 *   every other `.css` import was either invisible to esbuild
 *   (`prosemirror-view/style/prosemirror.css` was only in
 *   `layout.tsx`, which the bundle doesn't reach) or loaded as a
 *   string and discarded (`textLayer.css`). The visible symptoms
 *   were the XLSX cells stacking at one pixel (fixed last round),
 *   and DOCX/PDF missing ProseMirror's gap-cursor / atom-selection
 *   chrome and PDF's text-layer absolute positioning + transparent
 *   glyph rules required for selection.
 *
 *   This plugin closes the gap by making every `.css` import behave
 *   like Next's: the import becomes a side-effect that injects the
 *   stylesheet into `<head>` exactly once (deduped by a content
 *   hash) the first time the module loads. New `.css` imports added
 *   anywhere in the editor source graph in the future ship
 *   automatically — no second-source allow-list to keep in sync.
 *
 *   Implementation notes:
 *
 *   - Resolves the import in a custom `onResolve` so we can route to
 *     a "css-inject" namespace AHEAD of esbuild's externals check.
 *     Without this, `prosemirror-view/style/prosemirror.css` would
 *     match the `prosemirror-view` external prefix and never get
 *     inlined.
 *
 *   - The injected JS literal contains the CSS verbatim (no
 *     base64), prepends to `<head>` so host overrides cascade-win,
 *     and dedups by `id="officeai-css-<contentHash>"` so two
 *     entries that import the same sheet only inject once.
 *
 *   - Skips when `document` is undefined (SSR / Node), so the
 *     bundle stays usable in non-browser hosts.
 */
const cssInjectPlugin = {
  name: "officeai-css-inject",
  setup(build) {
    build.onResolve({ filter: /\.css$/ }, (args) => {
      // We deliberately do NOT use `build.resolve` here. Several of
      // our externals (`prosemirror-view`, `pdfjs-dist`) match by
      // prefix, so a `build.resolve("prosemirror-view/style/prosemirror.css")`
      // call gets short-circuited by esbuild's externals dispatch
      // and returns the bare specifier instead of an absolute path.
      // Resolve via Node's own algorithm rooted at the importer's
      // directory — that gets us a real file path even for CSS that
      // lives inside an externalised package, without altering the
      // JS-side externals contract.
      let absPath;
      if (path.isAbsolute(args.path)) {
        absPath = args.path;
      } else if (args.path.startsWith(".") || args.path.startsWith("/")) {
        absPath = path.resolve(args.resolveDir, args.path);
      } else {
        // Bare specifier (e.g. `prosemirror-view/style/prosemirror.css`).
        // Anchor the resolver at the importer's directory so pnpm's
        // nested layout finds the right copy.
        const requireFromImporter = createRequire(path.join(args.resolveDir, "package.json"));
        absPath = requireFromImporter.resolve(args.path);
      }
      return { path: absPath, namespace: "css-inject" };
    });

    build.onLoad({ filter: /.*/, namespace: "css-inject" }, (args) => {
      const css = readFileSync(args.path, "utf8");
      const hash = createHash("sha256").update(args.path).update("\0").update(css).digest("hex").slice(0, 16);
      const id = `officeai-css-${hash}`;
      // Self-contained: no shared helper module, no funky import
      // path resolution. Each .css import becomes a stand-alone
      // side-effect module the size of `${css.length}` plus ~300 B
      // of injection scaffolding. esbuild's tree-shaker can't drop
      // it (the import is bare), so it always runs at least once
      // per top-level entry that pulls it in.
      const contents = `// Auto-generated by build.mjs (officeai-css-inject plugin).
// Source: ${path.relative(REPO_ROOT, args.path)}
const STYLE_ID = ${JSON.stringify(id)};
if (typeof document !== "undefined" && !document.getElementById(STYLE_ID)) {
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.setAttribute("data-officeai-css", "1");
  style.textContent = ${JSON.stringify(css)};
  // Prepend so host stylesheets (loaded later via the host's own
  // <link>/<style>) cascade-win and can theme the editor without
  // fighting our defaults.
  if (document.head.firstChild) {
    document.head.insertBefore(style, document.head.firstChild);
  } else {
    document.head.appendChild(style);
  }
}
`;
      return { contents, loader: "js", resolveDir: path.dirname(args.path) };
    });
  },
};

const sharedOptions = {
  bundle: true,
  format: "esm",
  platform: "browser",
  target: "es2022",
  jsx: "automatic",
  external: EXTERNAL,
  alias: {
    // `@/foo/bar` in editor source → `apps/web/app/foo/bar`
    "@": APPS_WEB_APP,
    // strip the lone Next coupling so the bundle is host-agnostic
    "next/link": NEXT_LINK_SHIM,
  },
  plugins: [cssInjectPlugin],
  loader: {
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "js",
    ".jsx": "jsx",
    // `.css` is intentionally NOT registered here — the
    // `officeai-css-inject` plugin above intercepts every CSS
    // import in `onResolve` and routes it through the `css-inject`
    // namespace, which emits a self-injecting JS module. The
    // plugin runs before esbuild's loader-by-extension dispatch,
    // so any CSS reachable from the bundle's graph (the editor
    // source, plus the explicit `import "..."` lines we add to the
    // component entries) lands in the consumer's `<head>` on first
    // load — same behaviour as Next's automatic CSS handling, with
    // zero hand-maintained allow-lists.
  },
  // `"use client"` is a Next directive; safe to drop in non-Next hosts.
  // esbuild warns about unknown directives by default — silence it.
  logOverride: {
    "unsupported-jsx-comment": "silent",
    "ignored-bare-import": "silent",
  },
  // Components are split into per-format chunks via the `entryPoints`
  // map below. esbuild's code-splitting hoists shared code (the
  // `lib/shell` subtree, etc.) into a `chunks/` directory that all
  // four entries share at runtime.
  splitting: true,
  outdir: DIST,
  metafile: true,
};

// Phase-1 entries: small, hand-written in src/. Tiny enough to bundle
// in the same esbuild pass — that way `dist/` is the single output
// directory the package.json `exports` map points at.
const PHASE1_ENTRIES = {
  index: path.join(SRC, "index.ts"),
  "blanks/index": path.join(SRC, "blanks", "index.ts"),
  "blanks/docx": path.join(SRC, "blanks", "docx.ts"),
  "blanks/xlsx": path.join(SRC, "blanks", "xlsx.ts"),
  "blanks/pptx": path.join(SRC, "blanks", "pptx.ts"),
  "blanks/pdf": path.join(SRC, "blanks", "pdf.ts"),
  contract: path.join(SRC, "contract.ts"),
  mime: path.join(SRC, "mime.ts"),
};

// Phase-1.5 entries — the bundled editors.
const PHASE15_ENTRIES = {
  "components/docx": path.join(SRC, "components", "docx.ts"),
  "components/xlsx": path.join(SRC, "components", "xlsx.ts"),
  "components/pptx": path.join(SRC, "components", "pptx.ts"),
  "components/pdf": path.join(SRC, "components", "pdf.ts"),
};

const result = await esbuild.build({
  ...sharedOptions,
  entryPoints: { ...PHASE1_ENTRIES, ...PHASE15_ENTRIES },
});

if (result.warnings.length > 0) {
  for (const w of result.warnings) {
    console.warn(`[react-editors:build] warn: ${w.text}`);
  }
}

const jsOutputs = Object.keys(result.metafile.outputs).filter((p) => p.endsWith(".js"));
const cssOutputs = Object.keys(result.metafile.outputs).filter((p) => p.endsWith(".css"));
console.log(`[react-editors:build] emitted ${jsOutputs.length} JS files:`);
for (const p of jsOutputs.sort()) {
  console.log(`  - ${path.relative(PKG_ROOT, p)}`);
}
console.log(`[react-editors:build] emitted ${cssOutputs.length} CSS files:`);
for (const p of cssOutputs.sort()) {
  console.log(`  - ${path.relative(PKG_ROOT, p)}`);
}

// Copy hand-rolled .d.ts shims for the bundled editor components.
// `tsc` only emits types for the Phase-1 entries (see
// `tsconfig.types.json`) because typing the deeply-aliased
// `apps/web/app/...` graph would either drag the whole apps/web
// tsconfig into the package build (slow and brittle) or require
// `tsc` to traverse the editor source itself (it would). The
// hand-written declarations under `dist-types-template/components`
// expose only the embedded contract — that's the API we promise
// hosts, so it doubles as the public surface.
const TYPES_TEMPLATE = path.join(PKG_ROOT, "dist-types-template");
const TYPES_OUT = path.join(DIST, "components");
mkdirSync(TYPES_OUT, { recursive: true });
let copied = 0;
for (const entry of readdirSync(path.join(TYPES_TEMPLATE, "components"))) {
  if (!entry.endsWith(".d.ts")) continue;
  copyFileSync(path.join(TYPES_TEMPLATE, "components", entry), path.join(TYPES_OUT, entry));
  copied += 1;
}
console.log(`[react-editors:build] copied ${copied} .d.ts shim(s) into dist/components/.`);
