#!/usr/bin/env node
/**
 * License-graph scanner.
 *
 * Walks the resolved pnpm dependency graph (every entry in
 * `node_modules/.pnpm/<name@version>/node_modules/<name>/package.json`),
 * reads each package's declared SPDX `license` field, and:
 *
 *   - HARD FAILS on copyleft / commercial-incompatible licenses we have
 *     committed to never ship as a runtime dep (AGPL, GPL-only without an
 *     exception, SSPL, BUSL — see prompt.md §"Open-Source / Original Work").
 *   - WARNS on LGPL or "GPL with a classpath/library exception" entries
 *     (those are fine for our linkage model but worth surfacing).
 *   - WARNS on missing / unknown license fields (CI surface only — does NOT
 *     fail; the noise of upstream packages without a license field would
 *     gate every PR otherwise).
 *
 * No network calls. SPDX data comes from each package.json `license` /
 * `licenses` field. Unknown variants default to "warn".
 *
 * Self-check: when invoked with `--inject-agpl`, an artificial AGPL-3.0-only
 * entry is added to the scan list so test harnesses can verify the failing
 * path without a real banned package being present in the workspace.
 *
 * Run via `make licenses` or `node scripts/license-scan.mjs`.
 */

import { readdirSync, readFileSync, statSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "..");
const pnpmStore = resolve(root, "node_modules/.pnpm");

// SPDX expressions that must hard-fail. Matching is normalized (case-
// insensitive, trimmed). Entries are checked against the full SPDX expression
// so e.g. "MIT OR GPL-3.0-or-later" is OK (the dual-license out gives us MIT)
// while "GPL-3.0-only" by itself is NOT.
const BANNED_SPDX = new Set([
  "AGPL-1.0",
  "AGPL-1.0-only",
  "AGPL-1.0-or-later",
  "AGPL-3.0",
  "AGPL-3.0-only",
  "AGPL-3.0-or-later",
  "GPL-2.0-only",
  "GPL-3.0-only",
  "SSPL-1.0",
  "BUSL-1.1",
]);

// Surface but don't fail.
const WARN_SPDX_PREFIXES = ["LGPL-", "GPL-2.0-or-later", "GPL-3.0-or-later"];
const WARN_PHRASES = ["with classpath", "with library", "with linking", "with autoconf"];

const KNOWN_PERMISSIVE_PREFIXES = [
  "MIT",
  "ISC",
  "Apache-",
  "BSD-",
  "0BSD",
  "CC0-",
  "CC-BY-",
  "Unlicense",
  "Python-",
  "Zlib",
  "WTFPL",
  "Artistic-",
  "BlueOak-",
  "CDDL-",
  "EPL-",
  "MPL-",
  "PostgreSQL",
  "OFL-",
];

function classifySpdx(raw) {
  if (raw === undefined || raw === null) {
    return { level: "warn", reason: "missing license field" };
  }
  const s = String(raw).trim();
  if (!s) return { level: "warn", reason: "empty license field" };

  // SPDX expressions can be of the form "(MIT OR Apache-2.0)" or
  // "MIT AND BSD-3-Clause". Split on top-level `OR`/`AND` and check each
  // term — `OR` lets us pick a permissive escape hatch; `AND` requires all
  // terms to be acceptable.
  const cleaned = s.replace(/[()]/g, " ").replace(/\s+/g, " ").trim();
  const orTerms = splitTop(cleaned, /\s+OR\s+/i);
  if (orTerms.length > 1) {
    // If ANY OR-term is acceptable (no banned, no warn), the whole expression
    // is acceptable. Pick the first acceptable one for reporting.
    let best = null;
    let warns = [];
    for (const term of orTerms) {
      const c = classifySpdx(term);
      if (c.level === "ok") return { level: "ok", reason: `OR-clause permissive: ${term}` };
      if (c.level === "warn") warns.push(c);
      if (!best || c.level !== "fail") best = c;
    }
    if (warns.length > 0) return warns[0];
    return best ?? { level: "warn", reason: `unrecognised OR expression: ${s}` };
  }
  const andTerms = splitTop(cleaned, /\s+AND\s+/i);
  if (andTerms.length > 1) {
    let highest = { level: "ok", reason: "AND expression all permissive" };
    for (const term of andTerms) {
      const c = classifySpdx(term);
      if (c.level === "fail") return c;
      if (c.level === "warn") highest = c;
    }
    return highest;
  }

  const term = cleaned;
  if (BANNED_SPDX.has(term) || BANNED_SPDX.has(term.toUpperCase())) {
    return { level: "fail", reason: `banned SPDX: ${term}` };
  }
  for (const prefix of WARN_SPDX_PREFIXES) {
    if (term.toUpperCase().startsWith(prefix.toUpperCase())) {
      return { level: "warn", reason: `weak-copyleft / GPL-or-later: ${term}` };
    }
  }
  for (const phrase of WARN_PHRASES) {
    if (term.toLowerCase().includes(phrase)) {
      return { level: "warn", reason: `GPL with linking exception: ${term}` };
    }
  }
  for (const prefix of KNOWN_PERMISSIVE_PREFIXES) {
    if (term.toUpperCase().startsWith(prefix.toUpperCase())) return { level: "ok", reason: term };
  }
  return { level: "warn", reason: `unrecognised SPDX: ${term}` };
}

function splitTop(s, delim) {
  // The license expressions we encounter are simple enough that the top-level
  // OR/AND split is a plain regex; we strip parentheses upstream so we don't
  // need a real expression parser.
  return s.split(delim).map((t) => t.trim()).filter(Boolean);
}

function readLicenseField(pkg) {
  if (!pkg) return undefined;
  if (typeof pkg.license === "string") return pkg.license;
  if (pkg.license && typeof pkg.license === "object" && typeof pkg.license.type === "string") {
    return pkg.license.type;
  }
  if (Array.isArray(pkg.licenses) && pkg.licenses.length > 0) {
    const types = pkg.licenses.map((l) => (typeof l === "string" ? l : l?.type)).filter(Boolean);
    if (types.length > 0) return `(${types.join(" OR ")})`;
  }
  return undefined;
}

function isDir(p) {
  try {
    return statSync(p).isDirectory();
  } catch {
    return false;
  }
}

function discoverEntries() {
  // pnpm store layout: node_modules/.pnpm/<pkg-spec>/node_modules/<pkg>/package.json
  // For scoped packages, the inner dir is `@scope/name`. We walk one extra
  // level when we hit an `@scope` dir.
  const entries = [];
  if (!isDir(pnpmStore)) {
    console.warn(`warning: ${pnpmStore} not found — run \`pnpm install\` first.`);
    return entries;
  }
  for (const specDir of readdirSync(pnpmStore)) {
    if (specDir.startsWith(".")) continue;
    const inner = join(pnpmStore, specDir, "node_modules");
    if (!isDir(inner)) continue;
    for (const pkgName of readdirSync(inner)) {
      if (pkgName.startsWith(".")) continue;
      if (pkgName.startsWith("@")) {
        const scopeDir = join(inner, pkgName);
        if (!isDir(scopeDir)) continue;
        for (const sub of readdirSync(scopeDir)) {
          tryPushEntry(entries, join(scopeDir, sub));
        }
      } else {
        tryPushEntry(entries, join(inner, pkgName));
      }
    }
  }
  return entries;
}

function tryPushEntry(out, pkgDir) {
  const manifest = join(pkgDir, "package.json");
  try {
    const pkg = JSON.parse(readFileSync(manifest, "utf8"));
    if (!pkg.name) return;
    out.push({
      name: pkg.name,
      version: pkg.version ?? "0.0.0",
      license: readLicenseField(pkg),
    });
  } catch {
    // Skip malformed / missing manifests silently — pnpm's store has scratch
    // dirs we don't care about.
  }
}

function dedupe(entries) {
  // Multiple `<name@version+suffix>` dirs can resolve to the same logical
  // package; collapse by `name@version` for the report.
  const map = new Map();
  for (const e of entries) {
    const key = `${e.name}@${e.version}`;
    if (!map.has(key)) map.set(key, e);
  }
  return [...map.values()].sort((a, b) => a.name.localeCompare(b.name));
}

function scan(entries) {
  const fails = [];
  const warns = [];
  const okPermissive = new Map();
  for (const e of entries) {
    const c = classifySpdx(e.license);
    if (c.level === "fail") {
      fails.push({ ...e, reason: c.reason });
    } else if (c.level === "warn") {
      warns.push({ ...e, reason: c.reason });
    } else {
      const k = c.reason.replace(/^OR-clause permissive: /, "");
      okPermissive.set(k, (okPermissive.get(k) ?? 0) + 1);
    }
  }
  return { fails, warns, okPermissive };
}

function printSummary(entries, result) {
  console.log(`license-scan: scanned ${entries.length} unique packages\n`);
  console.log("| License (top-level term) | Count |");
  console.log("| ------------------------ | ----- |");
  const sorted = [...result.okPermissive.entries()].sort((a, b) => b[1] - a[1]);
  for (const [k, v] of sorted) {
    console.log(`| ${k.padEnd(24)} | ${String(v).padStart(5)} |`);
  }
  if (result.warns.length > 0) {
    console.log(`\n⚠ ${result.warns.length} warning(s):`);
    for (const w of result.warns) {
      console.log(`  - ${w.name}@${w.version}: ${w.reason}`);
    }
  }
  if (result.fails.length > 0) {
    console.log(`\n❌ ${result.fails.length} BANNED license(s):`);
    for (const f of result.fails) {
      console.log(`  - ${f.name}@${f.version}: ${f.reason}`);
    }
  }
}

function main() {
  const args = new Set(process.argv.slice(2));
  const injectAgpl = args.has("--inject-agpl");

  const entries = dedupe(discoverEntries());
  if (injectAgpl) {
    entries.push({
      name: "@officeai/__synthetic-agpl-test__",
      version: "0.0.0",
      license: "AGPL-3.0-only",
    });
  }

  const result = scan(entries);
  printSummary(entries, result);

  if (result.fails.length > 0) {
    console.error(
      `\nlicense-scan: FAILED — ${result.fails.length} disallowed license(s) in the dependency graph.`
    );
    return 1;
  }
  console.log("\n✅ license-scan: no banned licenses in the dependency graph.");
  return 0;
}

process.exit(main());
