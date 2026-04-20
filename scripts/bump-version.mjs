#!/usr/bin/env node
/**
 * Bump every publishable workspace package + the root manifest to a new
 * version in lockstep. Used by the auto-release workflow after CI is
 * green.
 *
 * Usage: node scripts/bump-version.mjs <new-version>
 *
 * The list of publishable packages is the canonical npm bundle that the
 * hof-os agent sandbox image consumes via `npm install -g
 * @officeai/agent`. Anything not in this list (`apps/web`, UI-only
 * packages, integration test harness) stays out of the bump because it
 * is intentionally `private: true`.
 */
import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");

// Mirrors scripts/check-architecture.mjs ALLOWED_INTERNAL_DEPS keys, minus
// the private-only packages (web, ui, design-tokens, comments, realtime,
// integration-tests, realtime-server). Keep in sync with package.json
// `publishConfig.access: public` and the sandbox image manifest.
const PUBLISHABLE = [
  "core",
  "text-formatting",
  "docx",
  "xlsx",
  "pptx",
  "pdf",
  "pdf-edit",
  "pdf-forms",
  "pdf-ocr",
  "pdf-annotations",
  "pdf-engine",
  "agent",
];

function fail(msg) {
  console.error(`bump-version: ${msg}`);
  process.exit(1);
}

const newVersion = process.argv[2];
if (!newVersion) {
  fail("usage: node scripts/bump-version.mjs <new-version>");
}
if (!/^\d+\.\d+\.\d+(-[\w.]+)?$/.test(newVersion)) {
  fail(`"${newVersion}" is not a valid semver version (e.g. 1.2.3 or 1.2.3-rc.1)`);
}

function bumpPkgJson(pkgPath, label) {
  if (!existsSync(pkgPath)) fail(`missing ${label} at ${pkgPath}`);
  const pkg = JSON.parse(readFileSync(pkgPath, "utf8"));
  const previous = pkg.version;
  pkg.version = newVersion;
  writeFileSync(pkgPath, JSON.stringify(pkg, null, 2) + "\n");
  console.log(`  ${label}: ${previous} → ${newVersion}`);
}

console.log(`Bumping office-ai workspace to ${newVersion}`);
bumpPkgJson(resolve(ROOT, "package.json"), "(root) office-ai");
for (const dirName of PUBLISHABLE) {
  bumpPkgJson(resolve(ROOT, "packages", dirName, "package.json"), `@officeai/${dirName}`);
}
console.log("done");
