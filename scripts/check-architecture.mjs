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
import { join, resolve } from "node:path";
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
  "@officeai/ui": ["@officeai/design-tokens"],
  "@officeai/docx": ["@officeai/core"],
  "@officeai/xlsx": ["@officeai/core"],
  "@officeai/pptx": ["@officeai/core"],
  "@officeai/agent": [
    "@officeai/core",
    "@officeai/docx",
    "@officeai/xlsx",
    "@officeai/pptx",
  ],
  "@officeai/integration-tests": [
    "@officeai/core",
    "@officeai/docx",
    "@officeai/xlsx",
    "@officeai/pptx",
    "@officeai/agent",
  ],
  "@officeai/web": [
    "@officeai/core",
    "@officeai/docx",
    "@officeai/xlsx",
    "@officeai/pptx",
    "@officeai/agent",
    "@officeai/ui",
    "@officeai/design-tokens",
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
};

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
  return [
    ...Object.keys(pkg.dependencies ?? {}),
    ...Object.keys(pkg.peerDependencies ?? {}),
  ];
}

function relPath(p) {
  return p.startsWith(ROOT) ? p.slice(ROOT.length + 1) : p;
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
  }

  if (violations.length === 0) {
    console.log("architecture-check: OK");
    console.log("  packages checked:", pkgs.map((p) => p.name).sort().join(", "));
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
