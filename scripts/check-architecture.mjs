#!/usr/bin/env node
/**
 * Architecture / dependency-graph linter.
 *
 * Reads each workspace package's package.json and validates that its
 * `dependencies` (and `peerDependencies`) only point at packages that
 * the architecture allows. Backs up the import-level boundaries
 * declared in eslint.config.mjs at the package-manifest level so that
 * BOTH layers must agree.
 *
 * The contract is intentionally declarative: the ALLOWED_INTERNAL_DEPS
 * map below is the single source of truth.
 *
 * Run via: `pnpm architecture` or `node scripts/check-architecture.mjs`.
 */
import { readdirSync, readFileSync, statSync } from "node:fs";
import { extname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(fileURLToPath(import.meta.url), "..", "..");
// Two kinds of workspace entries:
//   - parent dirs whose CHILDREN are packages (`packages/*`, `apps/*`)
//   - leaf dirs that are themselves packages (`tests`)
// Mirrors pnpm-workspace.yaml.
const PARENT_DIRS = ["packages", "apps"];
const LEAF_DIRS = ["tests"];

/**
 * For each internal package, list the OTHER internal packages it is
 * allowed to depend on (direct dependency, not transitive). Anything
 * not listed here is forbidden.
 *
 * Architectural intent:
 *  - core              → leaf; depends on nothing internal.
 *  - design-tokens     → leaf; depends on nothing internal.
 *  - ui                → presentation only; design-tokens.
 *  - docx              → headless model layer; core only.
 *  - xlsx              → headless model layer; core only (mirrors docx).
 *  - pptx              → headless model layer; core only (mirrors docx). Optional React renderer entry point under ./renderer/react keeps `react` an OPTIONAL peer dep, guarded by src/headless-invariant.test.ts.
 *  - agent             → orchestration / CLI; core + docx + xlsx + pptx.
 *  - integration-tests → can depend on anything (it's the consumer).
 *  - web               → top of stack; can depend on anything.
 */
const ALLOWED_INTERNAL_DEPS = {
  "@officeai/core": [],
  "@officeai/design-tokens": [],
  "@officeai/text-formatting": [],
  "@officeai/comments": [],
  "@officeai/realtime": ["@officeai/core"],
  "@officeai/ui": [
    "@officeai/design-tokens",
    "@officeai/text-formatting",
    "@officeai/comments",
    "@officeai/realtime",
  ],
  // docx/pptx may depend on xlsx for embedded-spreadsheet support
  // (chart data parts in docx; native OLE-spreadsheet shapes in pptx).
  // The dep is one-way — xlsx never reaches back into docx/pptx.
  "@officeai/docx": ["@officeai/core", "@officeai/text-formatting", "@officeai/comments", "@officeai/xlsx"],
  "@officeai/xlsx": ["@officeai/core", "@officeai/text-formatting", "@officeai/comments"],
  "@officeai/pptx": ["@officeai/core", "@officeai/text-formatting", "@officeai/comments", "@officeai/xlsx"],
  "@officeai/pdf-engine": [],
  // pdf depends on pdf-annotations for the unified comment/annotation
  // surface exposed at the top-level pdf package.
  "@officeai/pdf": ["@officeai/core", "@officeai/pdf-engine", "@officeai/pdf-annotations"],
  "@officeai/pdf-edit": ["@officeai/pdf"],
  "@officeai/pdf-annotations": ["@officeai/pdf"],
  "@officeai/pdf-forms": ["@officeai/pdf"],
  "@officeai/pdf-ocr": ["@officeai/pdf-engine"],
  "@officeai/agent": [
    "@officeai/core",
    "@officeai/docx",
    "@officeai/xlsx",
    "@officeai/pptx",
    "@officeai/pdf",
    "@officeai/pdf-edit",
    "@officeai/pdf-annotations",
    "@officeai/pdf-forms",
    "@officeai/pdf-ocr",
  ],
  // react-editors is the host-facing embed surface. Phase 1 ships the
  // blank-file builders + MIME constants; Phase 1.5 ships the bundled
  // editor components under ./components/{docx,xlsx,pptx,pdf} (built
  // from apps/web sources via esbuild — see packages/react-editors/
  // build.mjs).
  //
  // All @officeai/* deps are declared explicitly because Vite hosts
  // that set `resolve.preserveSymlinks: true`
  // can't see through pnpm's nested .pnpm/ store — declaring them as
  // direct deps makes pnpm hoist a top-level symlink at
  // `react-editors/node_modules/@officeai/<name>` where the resolver
  // actually looks. The Phase-1.5 bundles externalize these packages
  // (so they're shared with the host) which would fail to resolve
  // without this hoist. Same story for the heavy third-party deps
  // (pdfjs-dist, prosemirror-*, jszip, yjs, y-websocket)
  // which sit in package.json `dependencies` for the same reason.
  "@officeai/react-editors": [
    "@officeai/docx",
    "@officeai/xlsx",
    "@officeai/pptx",
    "@officeai/pdf",
    "@officeai/core",
    "@officeai/pdf-engine",
    "@officeai/pdf-annotations",
    "@officeai/ui",
    "@officeai/comments",
    "@officeai/design-tokens",
    "@officeai/realtime",
    "@officeai/text-formatting",
  ],
  "@officeai/realtime-server": [],
  "@officeai/integration-tests": [
    "@officeai/core",
    "@officeai/docx",
    "@officeai/xlsx",
    "@officeai/pptx",
    "@officeai/pdf",
    "@officeai/pdf-edit",
    "@officeai/pdf-annotations",
    "@officeai/pdf-forms",
    "@officeai/pdf-ocr",
    "@officeai/agent",
  ],
  "@officeai/web": [
    "@officeai/core",
    "@officeai/docx",
    "@officeai/xlsx",
    "@officeai/pptx",
    "@officeai/pdf",
    "@officeai/pdf-engine",
    "@officeai/pdf-edit",
    "@officeai/pdf-annotations",
    "@officeai/pdf-forms",
    "@officeai/pdf-ocr",
    "@officeai/agent",
    "@officeai/react-editors",
    "@officeai/ui",
    "@officeai/design-tokens",
    "@officeai/text-formatting",
    "@officeai/comments",
    "@officeai/realtime",
  ],
};

/**
 * Banned external dependencies per package. Reinforces "core/docx/agent
 * are headless" — no UI framework leaks into model code.
 */
const FORBIDDEN_EXTERNAL_DEPS = {
  "@officeai/core": ["react", "react-dom", "next"],
  "@officeai/docx": ["react", "react-dom", "next"],
  "@officeai/xlsx": ["react", "react-dom", "next"],
  // pptx exposes optional React renderer components via `./renderer/react`
  // entry point — react is an OPTIONAL peer dep, runtime imports are guarded
  // to the renderer/react/ subtree (see src/headless-invariant.test.ts).
  "@officeai/agent": ["react", "react-dom", "next"],
  "@officeai/design-tokens": ["react", "react-dom", "next"],
  "@officeai/text-formatting": ["react", "react-dom", "next"],
  "@officeai/realtime": ["react", "react-dom", "next"],
  "@officeai/pdf": ["react", "react-dom", "next"],
  "@officeai/pdf-engine": ["react", "react-dom", "next"],
  "@officeai/pdf-edit": ["react", "react-dom", "next"],
  "@officeai/pdf-annotations": ["react", "react-dom", "next"],
  "@officeai/pdf-forms": ["react", "react-dom", "next"],
  "@officeai/pdf-ocr": ["react", "react-dom", "next"],
};

const HOST_COUPLING_IMPORTS = [
  "@officeai/web",
  "@officeai/react-editors",
  "@officeai/realtime-server",
  "@modelcontextprotocol/sdk",
  "next",
  "react",
  "react-dom",
];

const MODEL_PACKAGE_HOST_IMPORTS = HOST_COUPLING_IMPORTS.filter(
  (specifier) => specifier !== "react" && specifier !== "react-dom"
);

const FORBIDDEN_SOURCE_IMPORTS = {
  "@officeai/core": [
    "@officeai/agent",
    "@officeai/web",
    "@officeai/react-editors",
    "@officeai/realtime-server",
    "@modelcontextprotocol/sdk",
    "next",
    "react",
    "react-dom",
  ],
  "@officeai/docx": MODEL_PACKAGE_HOST_IMPORTS,
  "@officeai/xlsx": HOST_COUPLING_IMPORTS,
  "@officeai/pptx": MODEL_PACKAGE_HOST_IMPORTS,
  "@officeai/pdf": HOST_COUPLING_IMPORTS,
  "@officeai/pdf-engine": HOST_COUPLING_IMPORTS,
  "@officeai/pdf-edit": HOST_COUPLING_IMPORTS,
  "@officeai/pdf-annotations": HOST_COUPLING_IMPORTS,
  "@officeai/pdf-forms": HOST_COUPLING_IMPORTS,
  "@officeai/pdf-ocr": HOST_COUPLING_IMPORTS,
  "@officeai/agent": [
    "@officeai/web",
    "@officeai/react-editors",
    "@officeai/realtime-server",
    "next",
    "react",
    "react-dom",
  ],
};

const SONALOOP_DESIGN_CONSUMERS = new Set(["@officeai/web", "@officeai/ui", "@officeai/react-editors"]);
const MIGRATION_ONLY_DESIGN_DEPS = {
  "@officeai/design-tokens": new Set(["@officeai/ui", "@officeai/web", "@officeai/react-editors"]),
};
const FORBIDDEN_DESIGN_DEPS = new Set(["lucide-react"]);
const DESIGN_DEPENDENCY_SOURCE = "sonaloop-design";

const SOURCE_EXTENSIONS = new Set([".ts", ".tsx", ".js", ".jsx", ".mjs", ".cjs"]);
const SKIPPED_SOURCE_DIRS = new Set(["node_modules", "dist", ".next", "coverage", ".turbo"]);
const IMPORT_SPECIFIER_RE =
  /\b(?:import|export)\s+(?:type\s+)?(?:[^"'()]*?\s+from\s+)?["']([^"']+)["']|import\s*\(\s*["']([^"']+)["']\s*\)|require\s*\(\s*["']([^"']+)["']\s*\)/g;

function tryReadPkg(pkgDir) {
  const pkgJsonPath = join(pkgDir, "package.json");
  try {
    const pkg = JSON.parse(readFileSync(pkgJsonPath, "utf8"));
    if (!pkg.name) return null;
    return { name: pkg.name, dir: pkgDir, pkg };
  } catch {
    return null;
  }
}

function discoverPackages() {
  const out = [];
  for (const base of PARENT_DIRS) {
    const baseDir = join(ROOT, base);
    let entries;
    try {
      entries = readdirSync(baseDir);
    } catch {
      continue;
    }
    for (const entry of entries) {
      const pkgDir = join(baseDir, entry);
      try {
        if (!statSync(pkgDir).isDirectory()) continue;
      } catch {
        continue;
      }
      const found = tryReadPkg(pkgDir);
      if (found) out.push(found);
    }
  }
  for (const leaf of LEAF_DIRS) {
    const found = tryReadPkg(join(ROOT, leaf));
    if (found) out.push(found);
  }
  return out;
}

function depKeys(pkg) {
  return [...Object.keys(pkg.dependencies ?? {}), ...Object.keys(pkg.peerDependencies ?? {})];
}

function relPath(p) {
  return p.startsWith(ROOT) ? p.slice(ROOT.length + 1) : p;
}

function walkSourceFiles(dir, out = []) {
  let entries;
  try {
    entries = readdirSync(dir);
  } catch {
    return out;
  }
  for (const entry of entries) {
    if (SKIPPED_SOURCE_DIRS.has(entry)) continue;
    const path = join(dir, entry);
    let st;
    try {
      st = statSync(path);
    } catch {
      continue;
    }
    if (st.isDirectory()) {
      walkSourceFiles(path, out);
      continue;
    }
    if (st.isFile() && SOURCE_EXTENSIONS.has(extname(entry))) out.push(path);
  }
  return out;
}

function sourceImportSpecifiers(file) {
  const text = readFileSync(file, "utf8");
  const specs = [];
  IMPORT_SPECIFIER_RE.lastIndex = 0;
  let match;
  while ((match = IMPORT_SPECIFIER_RE.exec(text))) {
    specs.push(match[1] ?? match[2] ?? match[3]);
  }
  return specs;
}

function matchesForbiddenSourceImport(specifier, forbidden) {
  return specifier === forbidden || specifier.startsWith(`${forbidden}/`);
}

function matchesImport(specifier, packageName) {
  return specifier === packageName || specifier.startsWith(`${packageName}/`);
}

function main() {
  const pkgs = discoverPackages();
  const knownInternal = new Set(pkgs.map((p) => p.name));
  const violations = [];

  for (const { name, dir, pkg } of pkgs) {
    const allowedInternal = ALLOWED_INTERNAL_DEPS[name];
    if (allowedInternal === undefined) {
      violations.push({
        package: name,
        file: relPath(join(dir, "package.json")),
        message: `Package "${name}" is not declared in scripts/check-architecture.mjs. Add it to ALLOWED_INTERNAL_DEPS so the dep graph stays explicit.`,
      });
      continue;
    }
    const allowed = new Set(allowedInternal);
    const deps = depKeys(pkg);

    for (const dep of deps) {
      if (knownInternal.has(dep) && !allowed.has(dep)) {
        violations.push({
          package: name,
          file: relPath(join(dir, "package.json")),
          message: `Forbidden internal dependency: "${name}" must not depend on "${dep}". Allowed internal deps: [${[...allowed].join(", ") || "none"}].`,
        });
      }
    }

    for (const dep of deps) {
      if (dep.startsWith("@officeai/design-") && dep !== "@officeai/design-tokens") {
        violations.push({
          package: name,
          file: relPath(join(dir, "package.json")),
          message: `Forbidden local design-system package "${dep}". Shared Sonaloop app design must come from "${DESIGN_DEPENDENCY_SOURCE}"; "@officeai/design-tokens" is the only temporary migration shim.`,
        });
      }
      if (dep === DESIGN_DEPENDENCY_SOURCE && !SONALOOP_DESIGN_CONSUMERS.has(name)) {
        violations.push({
          package: name,
          file: relPath(join(dir, "package.json")),
          message: `Forbidden design dependency: only app/UI shell packages may depend on "${DESIGN_DEPENDENCY_SOURCE}". Allowed consumers: [${[...SONALOOP_DESIGN_CONSUMERS].join(", ")}].`,
        });
      }
      if (FORBIDDEN_DESIGN_DEPS.has(dep)) {
        violations.push({
          package: name,
          file: relPath(join(dir, "package.json")),
          message: `Forbidden design dependency "${dep}". Icons must flow through "@officeai/ui/sonaloop-icons" backed by "${DESIGN_DEPENDENCY_SOURCE}".`,
        });
      }
      const migrationAllowed = MIGRATION_ONLY_DESIGN_DEPS[dep];
      if (migrationAllowed && !migrationAllowed.has(name)) {
        violations.push({
          package: name,
          file: relPath(join(dir, "package.json")),
          message: `Migration-only design dependency "${dep}" is only allowed in [${[...migrationAllowed].join(", ")}]. New design-system work must use "${DESIGN_DEPENDENCY_SOURCE}".`,
        });
      }
    }

    const forbidden = FORBIDDEN_EXTERNAL_DEPS[name] ?? [];
    for (const dep of deps) {
      if (forbidden.includes(dep)) {
        violations.push({
          package: name,
          file: relPath(join(dir, "package.json")),
          message: `Banned external dependency: "${name}" must not depend on "${dep}" (architectural boundary).`,
        });
      }
    }

    const forbiddenSourceImports = FORBIDDEN_SOURCE_IMPORTS[name] ?? [];
    if (forbiddenSourceImports.length > 0) {
      for (const file of walkSourceFiles(dir)) {
        for (const specifier of sourceImportSpecifiers(file)) {
          const forbiddenImport = forbiddenSourceImports.find((candidate) =>
            matchesForbiddenSourceImport(specifier, candidate)
          );
          if (!forbiddenImport) continue;
          violations.push({
            package: name,
            file: relPath(file),
            message: `Forbidden source import: "${name}" must not import "${specifier}". Host integrations belong behind the core adapter ports, not inside this package.`,
          });
        }
      }
    }

    for (const file of walkSourceFiles(dir)) {
      for (const specifier of sourceImportSpecifiers(file)) {
        if (matchesImport(specifier, "lucide-react")) {
          violations.push({
            package: name,
            file: relPath(file),
            message: `Forbidden source import: use "@officeai/ui/sonaloop-icons" instead of "lucide-react".`,
          });
        }
        if (!matchesImport(specifier, DESIGN_DEPENDENCY_SOURCE)) continue;
        if (SONALOOP_DESIGN_CONSUMERS.has(name)) continue;
        violations.push({
          package: name,
          file: relPath(file),
          message: `Forbidden source import: "${DESIGN_DEPENDENCY_SOURCE}" belongs only in app/UI shell packages, not in model, agent, runtime or test packages.`,
        });
      }
    }
  }

  if (violations.length === 0) {
    console.log("architecture-check: OK");
    console.log(
      "  packages checked:",
      pkgs
        .map((p) => p.name)
        .sort()
        .join(", ")
    );
    return 0;
  }

  console.error("architecture-check: FAILED\n");
  for (const v of violations) {
    console.error(`  ✖ [${v.package}] ${v.message}`);
    console.error(`      at ${v.file}\n`);
  }
  console.error(`Total violations: ${violations.length}`);
  return 1;
}

process.exit(main());
