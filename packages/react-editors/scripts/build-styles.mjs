#!/usr/bin/env node
/**
 * Derive `dist/styles.css` from `apps/web/app/globals.css`.
 *
 * Why a transform instead of a hand-maintained file:
 *
 *   The editors live in `apps/web/app/` and reference dozens of
 *   class selectors (`.xlsx-grid-cell`, `.prose-pm .ProseMirror`,
 *   `.pm-page-cap-top`, `.pptx-comment-flash`, …) along with CSS
 *   variables (`--divider`, `--ai-violet`, `--background`, …) that
 *   the host page doesn't know about. Without those rules the
 *   editor JSX renders but every cell collapses to `position:
 *   static` and the whole grid stacks at one pixel — exactly the
 *   "all cells overlap" bug we hit in hof-os.
 *
 *   apps/web/app/globals.css already owns the canonical version of
 *   all of these. Copying-and-trimming it at build time keeps the
 *   styles bit-for-bit in sync with what `make dev` renders inside
 *   apps/web, with zero hand-syncing risk.
 *
 * Transform rules (applied in order):
 *
 *   1. drop `@import "tailwindcss";`         (host owns Tailwind)
 *   2. drop `@source "...";`                 (Tailwind v4 scanner)
 *   3. drop the entire `@theme { ... }` /
 *      `@theme inline { ... }` blocks        (host owns design tokens)
 *   4. drop the global `html { ... }`,
 *      `body { ... }`, `::selection`,
 *      `::-webkit-scrollbar*`, the bare
 *      `a {…}` link transition,
 *      `:focus-visible {…}`, the
 *      `input::placeholder, textarea::placeholder`
 *      pair                                   (would clobber host page chrome)
 *
 *   Everything else — the brand-token `:root`/`.dark` blocks plus
 *   every `@keyframes` / namespaced editor class — is preserved
 *   verbatim and emitted to `dist/styles.css`.
 */

import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const PKG_ROOT = path.resolve(__dirname, "..");
const REPO_ROOT = path.resolve(PKG_ROOT, "..", "..");
const SRC_GLOBALS = path.join(REPO_ROOT, "apps", "web", "app", "globals.css");
const DIST = path.join(PKG_ROOT, "dist");
const OUT_PUBLIC = path.join(DIST, "styles.css");
// Mirror written into `src/` so esbuild can `import` it from each
// component entry (esbuild won't reach across to `dist/`). Marked
// `.generated.css` and gitignored — globals.css is the source of
// truth.
const OUT_SRC = path.join(PKG_ROOT, "src", "styles.generated.css");

/**
 * Walk a CSS source byte-by-byte and remove top-level rule blocks
 * (or `@xxx { ... }` blocks) whose selector header normalises to
 * any string in `selectorsToDrop`. Comments and string literals are
 * skipped so a `"}"` inside a content: declaration never confuses
 * the brace counter.
 */
function dropTopLevelBlocks(css, selectorsToDrop) {
  const drop = new Set(selectorsToDrop.map((s) => s.replace(/\s+/g, " ").trim()));
  const out = [];
  const n = css.length;
  let i = 0;

  /** Skip a /* ... *​/ comment starting at `pos`. Returns the index past it. */
  function skipComment(pos) {
    const end = css.indexOf("*/", pos + 2);
    return end < 0 ? n : end + 2;
  }
  /** Skip a "..." or '...' string starting at `pos`. Returns the index past the closing quote. */
  function skipString(pos) {
    const quote = css[pos];
    let p = pos + 1;
    while (p < n) {
      const ch = css[p];
      if (ch === "\\") {
        p += 2;
        continue;
      }
      if (ch === quote) return p + 1;
      p += 1;
    }
    return n;
  }

  while (i < n) {
    // Read header up to the next top-level `{` (which opens a rule
    // block) or top-level `;` (which terminates an at-rule statement
    // like `@import`/`@charset`/`@source`). Both end the chunk.
    const headerStart = i;
    let j = i;
    let kind = null; // "block" | "stmt"
    while (j < n) {
      const ch = css[j];
      if (ch === "/" && css[j + 1] === "*") {
        j = skipComment(j);
        continue;
      }
      if (ch === '"' || ch === "'") {
        j = skipString(j);
        continue;
      }
      if (ch === "{") {
        kind = "block";
        break;
      }
      if (ch === ";") {
        kind = "stmt";
        break;
      }
      j += 1;
    }
    if (j >= n) {
      out.push(css.slice(headerStart));
      break;
    }
    // Strip /* ... */ comments out of the captured header text before
    // normalising whitespace; the source uses block comments to label
    // sections (`/* ── Base styles ── */ html { … }`) and we don't
    // want those comments to defeat the selector predicate.
    const header = css
      .slice(headerStart, j)
      .replace(/\/\*[\s\S]*?\*\//g, "")
      .replace(/\s+/g, " ")
      .trim();
    if (kind === "stmt") {
      const stmtEnd = j + 1;
      if (drop.has(header) || (header.startsWith("@import") && /tailwindcss/.test(header)) || header.startsWith("@source")) {
        // skip
      } else {
        out.push(css.slice(headerStart, stmtEnd));
      }
      i = stmtEnd;
      continue;
    }
    // kind === "block": find matching closing brace.
    let k = j + 1;
    let depth = 1;
    while (k < n && depth > 0) {
      const ch = css[k];
      if (ch === "/" && css[k + 1] === "*") {
        k = skipComment(k);
        continue;
      }
      if (ch === '"' || ch === "'") {
        k = skipString(k);
        continue;
      }
      if (ch === "{") depth += 1;
      else if (ch === "}") depth -= 1;
      k += 1;
    }
    const blockEnd = k;
    if (drop.has(header)) {
      // skip
    } else {
      out.push(css.slice(headerStart, blockEnd));
    }
    i = blockEnd;
  }

  return out.join("");
}

const raw = readFileSync(SRC_GLOBALS, "utf8");

// Headers (normalised to single-space) we never want shipped to consumers.
// `dropTopLevelBlocks` also catches `@import "tailwindcss"` and `@source …`
// statements via dedicated startsWith checks for resilience.
const STRIP = [
  "@theme",
  "@theme inline",
  "html",
  "body",
  "::selection",
  "::-webkit-scrollbar",
  "::-webkit-scrollbar-track",
  "::-webkit-scrollbar-thumb",
  "::-webkit-scrollbar-thumb:hover",
  "a",
  ":focus-visible",
  "input::placeholder, textarea::placeholder",
];

const trimmed = dropTopLevelBlocks(raw, STRIP);

const banner = `/*\n * @officeai/react-editors — bundled editor styles.\n *\n * Auto-derived from apps/web/app/globals.css by\n * scripts/build-styles.mjs at build time. Do NOT edit this file by\n * hand — changes will be overwritten on the next build. Edit\n * globals.css instead.\n */\n\n`;

const finalCss = banner + trimmed.trim() + "\n";

mkdirSync(DIST, { recursive: true });
writeFileSync(OUT_PUBLIC, finalCss);
console.log(`[react-editors:build-styles] wrote ${path.relative(PKG_ROOT, OUT_PUBLIC)} (${finalCss.length.toLocaleString()} bytes)`);

// Also drop a copy into src/ so each component entry can `import` it
// and esbuild will emit per-entry CSS sidecars that Vite auto-loads.
writeFileSync(OUT_SRC, finalCss);
console.log(`[react-editors:build-styles] wrote ${path.relative(PKG_ROOT, OUT_SRC)} (${finalCss.length.toLocaleString()} bytes)`);
