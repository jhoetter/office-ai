#!/usr/bin/env node
/**
 * Copy PDF.js's CMap and standard-font assets from the workspace's
 * `pdfjs-dist` install into `apps/web/public/pdfjs/` so they ship as
 * static assets at `/pdfjs/cmaps/...` and `/pdfjs/standard_fonts/...`.
 *
 * Without these the PDF.js backend can't decode CJK/Arabic CMaps or
 * fall back to the 14 PDF base fonts when they aren't embedded —
 * either case manifests as silently empty or visibly garbled
 * selectable text in the viewer's text layer. See
 * `packages/pdf-engine/src/backends/pdfjs.ts` (`assetsBase`) and
 * `apps/web/app/pdf-viewer/PdfEditor.tsx` (`ensurePdfjsWorker`).
 *
 * Idempotent and skips writes when source and destination match by
 * size + mtime — fast on warm dev rebuilds.
 */
import { readdir, mkdir, copyFile, stat } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const webRoot = path.resolve(here, "..");
const repoRoot = path.resolve(webRoot, "..", "..");

async function resolvePdfjsRoot() {
  // `apps/web` depends on `pdfjs-dist` directly; `node_modules`
  // resolution goes through pnpm's `.pnpm` store. We resolve via
  // `package.json`'s metadata to avoid hard-coding the version.
  const candidates = [
    path.join(webRoot, "node_modules", "pdfjs-dist"),
    path.join(repoRoot, "node_modules", "pdfjs-dist"),
  ];
  for (const c of candidates) {
    if (existsSync(path.join(c, "package.json"))) return c;
  }
  // pnpm symlinks are often the actual install location.
  const pnpmDir = path.join(repoRoot, "node_modules", ".pnpm");
  if (existsSync(pnpmDir)) {
    const entries = await readdir(pnpmDir);
    const match = entries
      .filter((e) => e.startsWith("pdfjs-dist@"))
      .sort()
      .pop();
    if (match) {
      const candidate = path.join(pnpmDir, match, "node_modules", "pdfjs-dist");
      if (existsSync(path.join(candidate, "package.json"))) return candidate;
    }
  }
  throw new Error(
    "copy-pdfjs-assets: cannot locate pdfjs-dist install. " +
      "Run `pnpm install` first."
  );
}

async function copyDir(src, dst) {
  await mkdir(dst, { recursive: true });
  const entries = await readdir(src, { withFileTypes: true });
  for (const entry of entries) {
    const s = path.join(src, entry.name);
    const d = path.join(dst, entry.name);
    if (entry.isDirectory()) {
      await copyDir(s, d);
      continue;
    }
    const srcStat = await stat(s);
    let dstStat;
    try {
      dstStat = await stat(d);
    } catch {
      dstStat = null;
    }
    if (
      dstStat &&
      dstStat.size === srcStat.size &&
      dstStat.mtimeMs >= srcStat.mtimeMs
    ) {
      continue;
    }
    await copyFile(s, d);
  }
}

async function main() {
  const pdfjsRoot = await resolvePdfjsRoot();
  const targets = [
    { src: "cmaps", required: true },
    { src: "standard_fonts", required: true },
  ];
  const destRoot = path.join(webRoot, "public", "pdfjs");
  await mkdir(destRoot, { recursive: true });

  for (const t of targets) {
    const srcDir = path.join(pdfjsRoot, t.src);
    if (!existsSync(srcDir)) {
      if (t.required) {
        throw new Error(`copy-pdfjs-assets: missing ${srcDir}`);
      }
      continue;
    }
    const dstDir = path.join(destRoot, t.src);
    await copyDir(srcDir, dstDir);
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
