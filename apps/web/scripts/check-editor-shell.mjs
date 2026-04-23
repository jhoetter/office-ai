#!/usr/bin/env node
// CI gate: ensure every product editor wires through the shared
// EditorShell + design-tokens primitives. Failing this script keeps
// the three editors visually and structurally consistent.
//
// Checks:
//   1. Each editor TSX imports EditorShell from "../lib/shell".
//   2. Editor TSX files do not hard-code raw hex colours (must use
//      semantic tokens from @officeai/design-tokens). A small allowlist
//      handles legitimate exceptions (e.g. transparent SVG strokes).
//
// Run: pnpm --filter @officeai/web check:shell
import { readFile } from "node:fs/promises";
import { resolve, dirname, relative } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const APP_ROOT = resolve(__dirname, "..");

const EDITORS = [
  "app/editor/DocxEditor.tsx",
  "app/xlsx-editor/XlsxEditor.tsx",
  "app/pptx-editor/PptxEditor.tsx",
];

const EDITOR_DIRS_FOR_COLOR_CHECK = [
  "app/editor",
  "app/xlsx-editor",
  "app/pptx-editor",
];

// Hex literals we tolerate inside editor TSX files. Add sparingly and
// document why each entry is necessary. Comparison is lower-case, so
// every entry below MUST be lower-case too.
const HEX_ALLOWLIST = new Set([
  // DOCX theme picker palette in app/editor/DocxEditor.tsx —
  // these are *document content* fill/accent values that the user
  // applies to the rendered DOCX, not UI chrome. Putting them in
  // design-tokens would imply they round-trip through our theme
  // system, which they explicitly don't.
  "#ffffff",
  "#1f4e79",
  "#f4f6f8",
  "#334155",
  "#fff8f0",
  "#9a3412",
  "#f1f8f2",
  "#166534",
  // Transparency-indicator checker pattern in app/pptx-editor/
  // FillPicker.tsx — neutral grey checkerboard rendered inline as
  // SVG to flag alpha. No semantic colour exists for "this slot is
  // transparent".
  "#fff",
  "#ddd",
  "#dddddd",
]);

const errors = [];

async function checkEditorShellImport() {
  for (const rel of EDITORS) {
    const path = resolve(APP_ROOT, rel);
    const src = await readFile(path, "utf8");
    if (
      !/from\s+["']\.\.\/lib\/shell["']/.test(src) &&
      !/from\s+["']\.\/lib\/shell["']/.test(src) &&
      !/from\s+["']@\/lib\/shell["']/.test(src)
    ) {
      errors.push(`${rel}: missing import from "@/lib/shell" (or "../lib/shell")`);
      continue;
    }
    if (!/EditorShell/.test(src)) {
      errors.push(`${rel}: imports lib/shell but does not reference <EditorShell>`);
    }
  }
}

async function* walk(dir) {
  const { readdir } = await import("node:fs/promises");
  const entries = await readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    const full = resolve(dir, entry.name);
    if (entry.isDirectory()) {
      yield* walk(full);
    } else if (entry.isFile() && /\.(tsx|ts)$/.test(entry.name)) {
      // Test files legitimately use arbitrary hex strings as fixture
      // data (e.g. `authorColor: "#ff8800"` to assert round-trip
      // hydration). The colour gate is for runtime UI styling only.
      if (/\.test\.(tsx|ts)$/.test(entry.name)) continue;
      yield full;
    }
  }
}

const HEX_RE = /#[0-9a-fA-F]{3,8}\b/g;

async function checkNoHexColours() {
  for (const dirRel of EDITOR_DIRS_FOR_COLOR_CHECK) {
    const dirAbs = resolve(APP_ROOT, dirRel);
    for await (const file of walk(dirAbs)) {
      const src = await readFile(file, "utf8");
      const lines = src.split("\n");
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const m = line.match(HEX_RE);
        if (!m) continue;
        for (const hex of m) {
          if (HEX_ALLOWLIST.has(hex.toLowerCase())) continue;
          // Skip matches inside comments or string identifiers (cell
          // address `#REF!`, fragment refs, etc.). The simplest filter
          // is: only flag when the hex is preceded by `:` (CSS-style)
          // or the strings 'color'/'background'/'fill'/'stroke'.
          const lower = line.toLowerCase();
          const looksLikeColour =
            /:\s*"#[0-9a-f]/i.test(line) ||
            /background\w*\s*:\s*"#/.test(lower) ||
            /color\s*:\s*"#/.test(lower) ||
            /fill\s*[:=]\s*"#/.test(lower) ||
            /stroke\s*[:=]\s*"#/.test(lower) ||
            /borderColor\s*:\s*"#/.test(lower);
          if (!looksLikeColour) continue;
          errors.push(
            `${relative(APP_ROOT, file)}:${i + 1}: hard-coded colour ${hex} — use a token from @officeai/design-tokens`
          );
        }
      }
    }
  }
}

await checkEditorShellImport();
await checkNoHexColours();

if (errors.length > 0) {
  console.error("EditorShell / design-token gate failed:");
  for (const err of errors) console.error("  - " + err);
  console.error("");
  console.error(`Total violations: ${errors.length}`);
  process.exit(1);
}

console.log("EditorShell + design-tokens gate: OK");
