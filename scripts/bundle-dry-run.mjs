#!/usr/bin/env node
/**
 * Smoke-test the GitHub Releases bundles without uploading anything.
 *
 * Mirrors the `Build release bundle` step in
 * .github/workflows/auto-release.yml: runs `pnpm --filter <pkg>
 * --prod deploy <tmp>` for each tarball-shipped package, sanity-checks
 * the resulting bundle, prints its size, and cleans up.
 *
 * Tarballs covered:
 *   • @officeai/agent          — headless CLI sandbox image consumes
 *                                this. Must have a runnable
 *                                dist/cli.js with a node shebang.
 *   • @officeai/react-editors  — React component / blanks bundle that
 *                                embedding hosts pull into their web
 *                                apps. Must have dist/index.js +
 *                                dist/blanks/*.js.
 *
 * Used by `pnpm verify` so a broken `pnpm deploy` (workspace dep
 * cycle, missing build output, broken CLI shebang) is caught locally
 * before it reaches CI.
 *
 * Pure dry-run: nothing is uploaded, nothing is left on disk.
 */
import { execFileSync } from "node:child_process";
import {
  existsSync,
  lstatSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  symlinkSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, relative } from "node:path";

function run(cmd, args, opts = {}) {
  return execFileSync(cmd, args, { stdio: "inherit", ...opts });
}

function bytes(n) {
  const units = ["B", "KB", "MB", "GB"];
  let i = 0;
  let v = n;
  while (v >= 1024 && i < units.length - 1) {
    v /= 1024;
    i += 1;
  }
  return `${v.toFixed(1)} ${units[i]}`;
}

function dirSize(dir) {
  let total = 0;
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) total += dirSize(full);
    else if (entry.isFile()) total += statSync(full).size;
  }
  return total;
}

function statFile(path, label) {
  let stat;
  try {
    stat = statSync(path);
  } catch {
    throw new Error(`${label}: missing ${path} — did the build run first?`);
  }
  if (!stat.isFile()) {
    throw new Error(`${label}: ${path} is not a regular file`);
  }
}

function checkAgentBundle(out) {
  const cliPath = join(out, "dist", "cli.js");
  statFile(cliPath, "agent bundle");
  const firstLine = readFileSync(cliPath, "utf8").split("\n", 1)[0];
  if (!/^#!.*\bnode\b/.test(firstLine)) {
    throw new Error(`agent bundle: cli.js missing node shebang (got: ${JSON.stringify(firstLine)})`);
  }
  const nm = join(out, "node_modules");
  let nmStat;
  try {
    nmStat = statSync(nm);
  } catch {
    throw new Error("agent bundle: missing node_modules/ — workspace deps did not get inlined");
  }
  if (!nmStat.isDirectory()) {
    throw new Error("agent bundle: node_modules is not a directory");
  }
}

/**
 * Mirror the shameful-hoist step in `auto-release.yml`'s
 * `Build release bundle (officeai-react-editors)` job and the
 * fallback recommended for host postinstall checks. Required
 * because consumers running with `resolve.preserveSymlinks: true`
 * (common in Vite host apps) can't see through pnpm's `.pnpm/`
 * tree, so every transitive dep needs a top-level symlink at
 * `node_modules/<pkg>`. Keeping the dry-run in sync means any
 * regression in the deploy shape is caught locally before CI.
 */
function shamefullyHoist(deployRoot) {
  const nm = join(deployRoot, "node_modules");
  const pnpmDir = join(nm, ".pnpm");
  if (!existsSync(pnpmDir)) return 0;
  let hoisted = 0;
  const seen = new Set();
  for (const entry of readdirSync(pnpmDir)) {
    const inner = join(pnpmDir, entry, "node_modules");
    if (!existsSync(inner)) continue;
    for (const pkg of readdirSync(inner)) {
      if (pkg === ".bin") continue;
      if (pkg.startsWith("@")) {
        let subs;
        try {
          subs = readdirSync(join(inner, pkg));
        } catch {
          continue;
        }
        for (const sub of subs) {
          const full = `${pkg}/${sub}`;
          if (seen.has(full)) continue;
          seen.add(full);
          const target = join(nm, pkg, sub);
          if (lstatSync(target, { throwIfNoEntry: false })) continue;
          mkdirSync(dirname(target), { recursive: true });
          symlinkSync(relative(dirname(target), join(inner, pkg, sub)), target);
          hoisted += 1;
        }
      } else {
        if (seen.has(pkg)) continue;
        seen.add(pkg);
        const target = join(nm, pkg);
        if (lstatSync(target, { throwIfNoEntry: false })) continue;
        symlinkSync(relative(dirname(target), join(inner, pkg)), target);
        hoisted += 1;
      }
    }
  }
  return hoisted;
}

function checkReactEditorsBundle(out) {
  // Top-level barrel + each blanks entry point documented for hosts.
  // Catches a regressed exports map or a missing dist/.
  // The components/ entries are the Phase-1.5 bundled editors and
  // are required for the inline editor surface in embedding hosts.
  for (const rel of [
    join("dist", "index.js"),
    join("dist", "blanks", "index.js"),
    join("dist", "blanks", "docx.js"),
    join("dist", "blanks", "xlsx.js"),
    join("dist", "blanks", "pptx.js"),
    join("dist", "blanks", "pdf.js"),
    join("dist", "mime.js"),
    join("dist", "contract.js"),
    join("dist", "components", "docx.js"),
    join("dist", "components", "xlsx.js"),
    join("dist", "components", "pptx.js"),
    join("dist", "components", "pdf.js"),
    join("dist", "components", "docx.d.ts"),
    join("dist", "components", "xlsx.d.ts"),
    join("dist", "components", "pptx.d.ts"),
    join("dist", "components", "pdf.d.ts"),
    // Optional public stylesheet — `dist/styles.css` is the
    // human-readable extract emitted by `scripts/build-styles.mjs`
    // for hosts that want to load the editor CSS via a `<link>`
    // tag instead of relying on runtime injection. The bundled
    // components don't depend on it (they self-inject via the
    // `officeai-css-inject` plugin — see assertions below).
    join("dist", "styles.css"),
  ]) {
    statFile(join(out, rel), "react-editors bundle");
  }
  // Assert each bundled component ships the CSS it needs.
  //
  // The Next.js host gets editor CSS automatically via Next's
  // CSS-import discovery (see `apps/web/app/layout.tsx` and the
  // `import "*.css"` calls inside the editor source). Our esbuild
  // bundle gets the same coverage via the `officeai-css-inject`
  // plugin in `packages/react-editors/build.mjs`, which rewrites
  // every `.css` import into a self-injecting `<style>` tag. If a
  // future contributor breaks the plugin, removes a `.css`
  // side-effect import from a component entry, or upgrades a dep
  // in a way that makes a stylesheet vanish from the bundle, the
  // editors silently revert to the broken-styling failure mode the
  // user reported (xlsx cells stacked at one pixel, docx losing
  // ProseMirror selection chrome, pdf losing text-layer absolute
  // positioning). Catch that at release-build time instead.
  //
  // Search the assembled bundle (entry + transitive chunks) for a
  // distinctive selector from each required stylesheet. The
  // selectors are deliberately specific: matching `.ProseMirror`
  // alone would catch source-code references too.
  const assertCssShipped = (entryRel, label, needles) => {
    const entry = join(out, entryRel);
    const seen = new Set([entry]);
    const queue = [entry];
    let blob = "";
    while (queue.length > 0) {
      const file = queue.shift();
      const text = readFileSync(file, "utf8");
      blob += text;
      const importRe = /from\s+["']([^"']+\.js)["']/g;
      let m;
      while ((m = importRe.exec(text)) !== null) {
        const next = join(dirname(file), m[1]);
        if (!seen.has(next) && existsSync(next)) {
          seen.add(next);
          queue.push(next);
        }
      }
    }
    for (const needle of needles) {
      if (!blob.includes(needle)) {
        throw new Error(
          `react-editors bundle: ${label} entry did not ship "${needle}" — ` +
            `officeai-css-inject plugin or component .css imports regressed`
        );
      }
    }
  };
  // globals.css editor subset — XLSX cell positioning. Required by
  // every component (DOCX page sheets, PPTX comment flash, etc.
  // also rely on the same shared sheet via splitting).
  assertCssShipped("dist/components/xlsx.js", "XLSX", [".xlsx-grid-cell"]);
  assertCssShipped("dist/components/pptx.js", "PPTX", [".xlsx-grid-cell"]);
  // DOCX additionally needs prosemirror-view's runtime sheet
  // (selection / atom outline / hideselection caret). Without it
  // the editor still types but loses Word-style selection feedback.
  assertCssShipped("dist/components/docx.js", "DOCX", [".xlsx-grid-cell", ".ProseMirror-selectednode"]);
  // PDF additionally needs the text-layer stylesheet from
  // apps/web/app/pdf-viewer/textLayer.css (absolute positioning +
  // transparent glyphs required for selection over rendered pages).
  assertCssShipped("dist/components/pdf.js", "PDF", [".xlsx-grid-cell", ".officeai-pdf-text-layer"]);
  const nm = join(out, "node_modules");
  let nmStat;
  try {
    nmStat = statSync(nm);
  } catch {
    throw new Error("react-editors bundle: missing node_modules/ — workspace deps did not get inlined");
  }
  if (!nmStat.isDirectory()) {
    throw new Error("react-editors bundle: node_modules is not a directory");
  }
  // The four agent packages must be inlined; embedding hosts run the
  // bundle standalone with no access to office-ai's workspace.
  for (const dep of ["@officeai/docx", "@officeai/xlsx", "@officeai/pptx", "@officeai/pdf"]) {
    const depDir = join(nm, dep);
    let depStat;
    try {
      depStat = statSync(depDir);
    } catch {
      throw new Error(`react-editors bundle: missing inlined ${dep} under node_modules/`);
    }
    if (!depStat.isDirectory()) {
      throw new Error(`react-editors bundle: ${dep} is not a directory`);
    }
  }
  const hoisted = shamefullyHoist(out);
  console.log(`bundle:dry-run -> @officeai/react-editors shamefully-hoisted ${hoisted} package(s)`);
  // Spot-check transitive deps that the four agents AND the
  // bundled editor components reach for under preserveSymlinks: true.
  // Each must resolve from the deploy root.
  for (const dep of [
    "@officeai/core",
    "@officeai/pdf-engine",
    "@officeai/pdf-annotations",
    "@officeai/ui",
    "@officeai/comments",
    "@officeai/realtime",
    "@officeai/text-formatting",
    "@officeai/design-tokens",
    "pdf-lib",
    "fast-xml-parser",
    "js-sha256",
    "@e965/xlsx",
    // Phase-1.5 editor third-party externals — must be hoisted so the
    // bundled editor JS files can resolve them in the host's Vite.
    "pdfjs-dist",
    "lucide-react",
    "jszip",
    "yjs",
    "y-websocket",
    "prosemirror-view",
    "prosemirror-state",
    "prosemirror-model",
  ]) {
    const depDir = join(nm, dep);
    if (!existsSync(depDir)) {
      throw new Error(
        `react-editors bundle: transitive ${dep} not resolvable from deploy root after shameful-hoist`
      );
    }
  }
}

const TARGETS = [
  { pkg: "@officeai/agent", check: checkAgentBundle },
  { pkg: "@officeai/react-editors", check: checkReactEditorsBundle },
];

let exitCode = 0;
for (const { pkg, check } of TARGETS) {
  const out = mkdtempSync(join(tmpdir(), "officeai-bundle-dryrun-"));
  try {
    console.log(`bundle:dry-run -> ${pkg} into ${out}`);
    run("pnpm", ["--filter", pkg, "--prod", "deploy", out]);
    check(out);
    const total = dirSize(out);
    console.log(`bundle:dry-run OK — ${pkg}, total ${bytes(total)}`);
  } catch (err) {
    console.error(`bundle:dry-run FAILED for ${pkg}: ${err.message}`);
    exitCode = 1;
  } finally {
    rmSync(out, { recursive: true, force: true });
  }
}
process.exit(exitCode);
