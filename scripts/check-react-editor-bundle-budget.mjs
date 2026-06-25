#!/usr/bin/env node
/**
 * Enforce the embeddable editor bundle budgets.
 *
 * The gate measures what a host imports from
 * `@officeai/react-editors/components/<format>`: the entry file plus
 * statically imported shared chunks. Dynamic imports are reported
 * separately so optional collaboration code can stay lazy.
 */
import { existsSync, readFileSync, statSync } from "node:fs";
import { dirname, relative, resolve } from "node:path";
import { gzipSync } from "node:zlib";

const DIST = resolve("packages/react-editors/dist");
const FORMATS = ["docx", "xlsx", "pptx", "pdf"];

const BUDGETS = {
  docx: { raw: 1_100_000, gzip: 230_000 },
  xlsx: { raw: 1_350_000, gzip: 270_000 },
  pptx: { raw: 1_100_000, gzip: 230_000 },
  pdf: { raw: 820_000, gzip: 180_000 },
};

const STATIC_ALLOWED = {
  docx: new Set([
    "@officeai/core",
    "@officeai/docx",
    "@officeai/xlsx",
    "jszip",
    "prosemirror-state",
    "prosemirror-view",
    "react",
    "react-dom",
    "sonaloop-design",
  ]),
  xlsx: new Set(["@officeai/xlsx", "jszip", "react", "react-dom", "sonaloop-design"]),
  pptx: new Set(["@officeai/pptx", "@officeai/xlsx", "jszip", "react", "react-dom", "sonaloop-design"]),
  pdf: new Set([
    "@officeai/pdf",
    "@officeai/pdf-engine",
    "pdfjs-dist",
    "react",
    "react-dom",
    "sonaloop-design",
  ]),
};

const HEAVY_DEP_ALLOWLIST = new Map([
  ["jszip", "OOXML zip container read/write across DOCX/XLSX/PPTX."],
  ["pdfjs-dist", "PDF parser/renderer, only allowed in the PDF editor graph."],
  ["prosemirror-model", "DOCX rich-text model dependency."],
  ["prosemirror-state", "DOCX rich-text editor state dependency."],
  ["prosemirror-view", "DOCX rich-text view dependency."],
  ["yjs", "Optional realtime collaboration, must stay in the dynamic RoomClient chunk."],
  ["y-websocket", "Optional realtime collaboration transport, must stay in the dynamic RoomClient chunk."],
]);

const REALTIME_DYNAMIC_ALLOWED = new Set(["yjs", "y-websocket"]);
const IMPORT_RE = /(?:import|export)\s+(?:[^"']*?\s+from\s*)?["']([^"']+)["']|import\(["']([^"']+)["']\)/g;

let failed = false;

function markFailure(message) {
  failed = true;
  console.error(`bundle:budget FAILED: ${message}`);
}

function formatBytes(n) {
  const units = ["B", "KB", "MB"];
  let value = n;
  let i = 0;
  while (value >= 1024 && i < units.length - 1) {
    value /= 1024;
    i += 1;
  }
  return `${value.toFixed(1)} ${units[i]}`;
}

function packageName(specifier) {
  if (specifier.startsWith(".") || specifier.startsWith("/") || specifier.startsWith("node:")) {
    return null;
  }
  if (specifier.startsWith("@")) return specifier.split("/").slice(0, 2).join("/");
  return specifier.split("/")[0];
}

function resolveRelativeJs(fromFile, specifier) {
  if (!specifier.startsWith(".") || !specifier.endsWith(".js")) return null;
  return resolve(dirname(fromFile), specifier);
}

function collectStaticGraph(entry) {
  const seen = new Set();
  const queue = [entry];
  const staticBare = new Set();
  const dynamicImports = new Set();

  while (queue.length > 0) {
    const file = queue.shift();
    if (!file || seen.has(file)) continue;
    if (!existsSync(file)) {
      markFailure(`missing imported chunk ${relative(DIST, file)}`);
      continue;
    }
    seen.add(file);
    const source = readFileSync(file, "utf8");
    IMPORT_RE.lastIndex = 0;
    let m;
    while ((m = IMPORT_RE.exec(source)) !== null) {
      const specifier = m[1] ?? m[2];
      const isDynamic = Boolean(m[2]);
      const pkg = packageName(specifier);
      if (pkg && !isDynamic) staticBare.add(pkg);
      if (pkg && isDynamic) dynamicImports.add(specifier);

      const relativeImport = resolveRelativeJs(file, specifier);
      if (!relativeImport) continue;
      if (isDynamic) dynamicImports.add(relativeImport);
      else queue.push(relativeImport);
    }
  }

  return { files: seen, staticBare, dynamicImports };
}

function measureFiles(files) {
  const buffers = [...files].map((file) => readFileSync(file));
  const raw = buffers.reduce((sum, buf) => sum + buf.byteLength, 0);
  const gzip = gzipSync(Buffer.concat(buffers), { level: 9 }).byteLength;
  return { raw, gzip };
}

function collectBarePackages(file) {
  const packages = new Set();
  const source = readFileSync(file, "utf8");
  IMPORT_RE.lastIndex = 0;
  let m;
  while ((m = IMPORT_RE.exec(source)) !== null) {
    const pkg = packageName(m[1] ?? m[2]);
    if (pkg) packages.add(pkg);
  }
  return packages;
}

function checkPackageHeavyDeps() {
  const pkg = JSON.parse(readFileSync(resolve("packages/react-editors/package.json"), "utf8"));
  const deps = Object.keys(pkg.dependencies ?? {});
  const heavyDeps = deps.filter((dep) => HEAVY_DEP_ALLOWLIST.has(dep)).sort();
  const disallowedHeavy = deps.filter((dep) => {
    if (HEAVY_DEP_ALLOWLIST.has(dep)) return false;
    return (
      dep === "lucide-react" || dep === "xlsx" || dep.startsWith("pdfjs") || dep.startsWith("prosemirror-")
    );
  });

  console.log("bundle:budget heavy direct deps:");
  for (const dep of heavyDeps) {
    console.log(`  - ${dep}: ${HEAVY_DEP_ALLOWLIST.get(dep)}`);
  }
  if (disallowedHeavy.length > 0) {
    markFailure(`unexpected heavy direct dependency/dependencies: ${disallowedHeavy.join(", ")}`);
  }
}

checkPackageHeavyDeps();

for (const format of FORMATS) {
  const entry = resolve(DIST, "components", `${format}.js`);
  if (!existsSync(entry)) {
    markFailure(`missing ${relative(process.cwd(), entry)}; run pnpm --filter @officeai/react-editors build`);
    continue;
  }

  const graph = collectStaticGraph(entry);
  const size = measureFiles(graph.files);
  const budget = BUDGETS[format];
  const files = [...graph.files].map((file) => relative(DIST, file)).sort();
  const staticPackages = [...graph.staticBare].sort();

  console.log(
    `bundle:budget ${format}: ${files.length} file(s), raw ${formatBytes(size.raw)} / ${formatBytes(
      budget.raw
    )}, gzip ${formatBytes(size.gzip)} / ${formatBytes(budget.gzip)}`
  );
  console.log(`  static files: ${files.join(", ")}`);
  console.log(`  static externals: ${staticPackages.join(", ") || "(none)"}`);

  if (size.raw > budget.raw) {
    markFailure(`${format} raw ${size.raw} exceeds budget ${budget.raw}`);
  }
  if (size.gzip > budget.gzip) {
    markFailure(`${format} gzip ${size.gzip} exceeds budget ${budget.gzip}`);
  }

  const allowed = STATIC_ALLOWED[format];
  const unexpected = staticPackages.filter((pkg) => !allowed.has(pkg));
  if (unexpected.length > 0) {
    markFailure(`${format} has unexpected static external(s): ${unexpected.join(", ")}`);
  }
  for (const realtime of REALTIME_DYNAMIC_ALLOWED) {
    if (graph.staticBare.has(realtime)) {
      markFailure(`${format} statically imports ${realtime}; realtime must stay lazy`);
    }
  }

  const dynamicFiles = [...graph.dynamicImports].filter((specifier) => specifier.startsWith("/"));
  if (dynamicFiles.length > 0) {
    for (const dynamicFile of dynamicFiles.sort()) {
      const pkgs = collectBarePackages(dynamicFile);
      const dynamicSize = measureFiles(new Set([dynamicFile]));
      console.log(
        `  dynamic ${relative(DIST, dynamicFile)}: raw ${formatBytes(dynamicSize.raw)}, gzip ${formatBytes(
          dynamicSize.gzip
        )}, externals ${[...pkgs].sort().join(", ") || "(none)"}`
      );
      const unexpectedDynamic = [...pkgs].filter((pkg) => !REALTIME_DYNAMIC_ALLOWED.has(pkg));
      if (unexpectedDynamic.length > 0) {
        markFailure(
          `${format} dynamic ${relative(DIST, dynamicFile)} has unexpected external(s): ${unexpectedDynamic.join(
            ", "
          )}`
        );
      }
    }
  }
}

if (failed) process.exit(1);
console.log("bundle:budget OK");
