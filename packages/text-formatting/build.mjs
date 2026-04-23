#!/usr/bin/env node
// Bundle @officeai/text-formatting into a single self-contained ESM file.
//
// Background: this package is consumed two ways —
//   1. Bundlers (Vite, esbuild, Next) inside the workspace and inside
//      `@officeai/react-editors` (which inlines us) — they happily
//      handle raw `.ts` and were the historical reason the package
//      shipped `main: "./src/index.ts"`.
//   2. Raw Node ESM (e.g. `scripts/validate-ooxml-schemas.mjs` reaching
//      us transitively through `packages/{core,docx,xlsx,pptx}/dist/`)
//      — Node's resolver chokes on `.ts`, and `tsc`'s `moduleResolution:
//      bundler` output omits the `.js` extensions Node requires.
//
// esbuild bundling sidesteps both issues: one file, no relative imports,
// runs in Node and any bundler. Types are emitted by a sibling `tsc`
// pass driven by the package.json `build` script.

import esbuild from "esbuild";
import { mkdirSync, rmSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const __filename = fileURLToPath(import.meta.url);
const PKG_ROOT = path.dirname(__filename);
const DIST = path.join(PKG_ROOT, "dist");

rmSync(DIST, { recursive: true, force: true });
mkdirSync(DIST, { recursive: true });

await esbuild.build({
  entryPoints: [path.join(PKG_ROOT, "src", "index.ts")],
  outfile: path.join(DIST, "index.js"),
  bundle: true,
  format: "esm",
  platform: "neutral",
  target: "es2022",
  sourcemap: true,
});

console.log("[text-formatting:build] emitted dist/index.js");
