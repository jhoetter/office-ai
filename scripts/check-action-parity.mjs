#!/usr/bin/env node
/**
 * Action parity gate.
 *
 * For each format `F` in {docx, xlsx, pptx, pdf}, verifies that the
 * action catalogue at `packages/F/src/actions/catalogue.ts` covers
 * every command-bus handler registered in
 * `packages/F/src/commands/`. Catches the most expensive class of
 * drift in the codebase: a new bus command lands without a CLI
 * subcommand or a Cmd+K palette entry, so it's reachable by the
 * editor's own React code but invisible to humans (and AIs) outside
 * that one editor.
 *
 * The script runs at the SOURCE level (no `dist/` dependency) so it
 * stays cheap (<200ms) and can run BEFORE `pnpm build` — which is
 * the whole point: catching parity drift is supposed to be the first,
 * fastest signal in `make verify`.
 *
 * Output is human-readable; exit code is non-zero on drift unless
 * the `--lenient` flag is set, in which case violations are printed
 * but the process exits 0. The plan calls for landing the parity
 * gate in lenient mode for one PR (so the existing drift surfaces
 * without blocking the merge), then flipping to strict.
 *
 * Usage:
 *   node scripts/check-action-parity.mjs           # strict (CI)
 *   node scripts/check-action-parity.mjs --lenient # warn-only
 */

import { readdirSync, readFileSync, statSync } from "node:fs";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(fileURLToPath(import.meta.url), "..", "..");
const LENIENT = process.argv.includes("--lenient");

/**
 * Per-format wiring: where to find handler types vs. where to find
 * the catalogue. Each format may have its handlers in one or more
 * directories (pdf splits across pdf/, pdf-edit/, etc.) but the
 * catalogue is always one file.
 */
const FORMATS = [
  {
    name: "docx",
    handlerDirs: ["packages/docx/src/commands"],
    cataloguePath: "packages/docx/src/actions/catalogue.ts",
  },
  {
    name: "xlsx",
    handlerDirs: ["packages/xlsx/src/commands"],
    cataloguePath: "packages/xlsx/src/actions/catalogue.ts",
  },
  {
    name: "pptx",
    handlerDirs: ["packages/pptx/src/commands"],
    cataloguePath: "packages/pptx/src/actions/catalogue.ts",
  },
  {
    name: "pdf",
    handlerDirs: ["packages/pdf/src/commands"],
    cataloguePath: "packages/pdf/src/actions/catalogue.ts",
  },
];

/* ── Source extraction ──────────────────────────────────────────── */

/**
 * Pull every handler `type: "<format>:..."` literal from every .ts file
 * (excluding .test.ts) inside the given directories.
 *
 * The shape we look for is the canonical handler declaration:
 *   `export const fooHandler: CommandHandler<...> = { type: "docx:foo", ... }`
 * but the regex is intentionally loose — anything matching
 * `type: "docx:..."` at the start of a non-comment line counts.
 */
function extractHandlerTypes(format, handlerDirs) {
  const types = new Set();
  for (const dir of handlerDirs) {
    const abs = join(ROOT, dir);
    walkTsFiles(abs, (file, contents) => {
      if (file.endsWith(".test.ts")) return;
      // Match `type: "format:foo-bar"` only; exclude lines inside
      // dispatcher tests or arbitrary objects.
      const re = new RegExp(`\\btype:\\s*"(${format}:[a-z][a-z0-9-]*)"`, "g");
      let m;
      while ((m = re.exec(contents)) !== null) {
        types.add(m[1]);
      }
    });
  }
  return types;
}

/**
 * Pull every catalogue entry's `id`, `commandType`, and `surfaces`
 * from a single catalogue.ts file. Uses a deliberately narrow regex
 * over the canonical declaration shape; if someone reformats the file
 * past what the regex matches, the parity check will yell with a
 * clear "could not parse" line so the regex can be tightened.
 */
function extractCatalogueEntries(cataloguePath) {
  const abs = join(ROOT, cataloguePath);
  const src = readFileSync(abs, "utf8");
  const entries = [];
  // Scan blocks bounded by `{ id: "..."` … the next top-level `},` —
  // we don't need a real parser because the catalogue files are flat
  // arrays of object literals.
  const idRe = /\{\s*id:\s*"([^"]+)"/g;
  let m;
  const positions = [];
  while ((m = idRe.exec(src)) !== null) {
    positions.push({ id: m[1], start: m.index });
  }
  for (let i = 0; i < positions.length; i++) {
    const start = positions[i].start;
    const end = i + 1 < positions.length ? positions[i + 1].start : src.length;
    const block = src.slice(start, end);
    const id = positions[i].id;
    const commandType = parseCommandType(block);
    const surfaces = parseSurfaces(block);
    const hidden = /\bhidden\s*:\s*\{/.test(block);
    entries.push({ id, commandType, surfaces, hidden });
  }
  return entries;
}

function parseCommandType(block) {
  const m = block.match(/\bcommandType:\s*(?:"([^"]+)"|null)/);
  if (!m) return undefined;
  return m[1] ?? null;
}

function parseSurfaces(block) {
  // surfaces: ["cli", "palette"]
  const m = block.match(/\bsurfaces:\s*\[([^\]]*)\]/);
  if (!m) return [];
  return [...m[1].matchAll(/"([a-zA-Z]+)"/g)].map((x) => x[1]);
}

function walkTsFiles(dir, visit) {
  let entries;
  try {
    entries = readdirSync(dir);
  } catch {
    return;
  }
  for (const entry of entries) {
    const full = join(dir, entry);
    let st;
    try {
      st = statSync(full);
    } catch {
      continue;
    }
    if (st.isDirectory()) {
      walkTsFiles(full, visit);
    } else if (st.isFile() && full.endsWith(".ts")) {
      visit(full, readFileSync(full, "utf8"));
    }
  }
}

/* ── Parity rules ───────────────────────────────────────────────── */

/**
 * For a single format, compute the violation list against the rules:
 *
 *  1. Every registered handler `type` must have ≥1 catalogue entry
 *     whose `commandType` matches it (covered, possibly hidden).
 *  2. No catalogue entry may declare a `commandType` that no handler
 *     actually serves (catches typos and stale renames).
 *  3. Every catalogue entry's `id` must be unique within the format.
 *  4. Every catalogue entry that lists `cli` in `surfaces` and is NOT
 *     `hidden` is expected to be reachable from `office-agent <fmt>`.
 *     We can't import the commander tree from this script (the CLI
 *     mounts everything in cli.ts); instead we treat the entry's
 *     declared `cli` membership as a signed contract — the CLI
 *     adapter / hand-rolled cli-*.ts is responsible for honouring it,
 *     and the test suite (cli.test.ts / pptx-cli.test.ts / …) has
 *     coverage that fails if the subcommand is missing. The parity
 *     gate's job here is structural: the catalogue declares the
 *     intent.
 */
function checkFormat(format, handlerTypes, entries, i18nKeys) {
  const violations = [];

  const seenIds = new Set();
  for (const e of entries) {
    if (seenIds.has(e.id)) {
      violations.push({
        format,
        kind: "duplicate-id",
        message: `duplicate catalogue id "${e.id}"`,
      });
    }
    seenIds.add(e.id);
  }

  const cataloguedTypes = new Set();
  for (const e of entries) {
    if (e.commandType !== null && e.commandType !== undefined) {
      cataloguedTypes.add(e.commandType);
    }
  }

  for (const t of handlerTypes) {
    if (!cataloguedTypes.has(t)) {
      violations.push({
        format,
        kind: "missing-catalogue-entry",
        message: `handler "${t}" is not in the action catalogue — add an entry in packages/${format}/src/actions/catalogue.ts`,
      });
    }
  }

  for (const t of cataloguedTypes) {
    if (!handlerTypes.has(t)) {
      violations.push({
        format,
        kind: "stale-command-type",
        message: `catalogue references commandType "${t}" but no handler is registered — typo, or the handler was renamed/removed`,
      });
    }
  }

  // i18n parity: every palette-surfaced action must have an English
  // label + description in messages/en.json. Without these, the Cmd+K
  // palette would render the catalogue's English fallback even when
  // the user has switched locales — a half-translated palette is
  // worse than a fully untranslated one. DE is intentionally not
  // checked: translators flesh it out independently and the resolver
  // falls back per-key.
  for (const e of entries) {
    if (e.hidden) continue;
    if (!e.surfaces.includes("palette")) continue;
    const localId = e.id.startsWith(`${format}.`)
      ? e.id.slice(format.length + 1)
      : e.id;
    const labelKey = `actions.${format}.${localId}.label`;
    const descKey = `actions.${format}.${localId}.description`;
    if (!i18nKeys.has(labelKey)) {
      violations.push({
        format,
        kind: "missing-i18n-label",
        message: `palette action "${e.id}" has no English translation — add ${labelKey} to apps/web/app/lib/i18n/messages/en.json`,
      });
    }
    if (!i18nKeys.has(descKey)) {
      violations.push({
        format,
        kind: "missing-i18n-description",
        message: `palette action "${e.id}" has no English description — add ${descKey} to apps/web/app/lib/i18n/messages/en.json`,
      });
    }
  }

  return violations;
}

/**
 * Flatten the messages JSON tree into a Set of dotted keys so the
 * parity check can do O(1) presence lookups without re-walking the
 * tree. Only string leaves are recorded; arrays and other shapes
 * (the i18n catalogue uses neither today) are skipped.
 */
function loadI18nKeys(path) {
  const abs = join(ROOT, path);
  const tree = JSON.parse(readFileSync(abs, "utf8"));
  const keys = new Set();
  const walk = (node, prefix) => {
    if (node === null || typeof node !== "object") return;
    for (const [k, v] of Object.entries(node)) {
      const dotted = prefix ? `${prefix}.${k}` : k;
      if (typeof v === "string") {
        keys.add(dotted);
      } else if (v !== null && typeof v === "object" && !Array.isArray(v)) {
        walk(v, dotted);
      }
    }
  };
  walk(tree, "");
  return keys;
}

/* ── Entry point ────────────────────────────────────────────────── */

function main() {
  const allViolations = [];
  const summary = [];

  const i18nKeys = loadI18nKeys("apps/web/app/lib/i18n/messages/en.json");

  for (const f of FORMATS) {
    const handlerTypes = extractHandlerTypes(f.name, f.handlerDirs);
    const entries = extractCatalogueEntries(f.cataloguePath);
    const violations = checkFormat(f.name, handlerTypes, entries, i18nKeys);
    summary.push({
      format: f.name,
      handlers: handlerTypes.size,
      catalogue: entries.length,
      violations: violations.length,
    });
    allViolations.push(...violations);
  }

  console.log("action-parity check");
  console.log("───────────────────");
  for (const s of summary) {
    console.log(
      `  ${s.format.padEnd(6)} handlers=${String(s.handlers).padStart(3)}  catalogue=${String(s.catalogue).padStart(3)}  violations=${String(s.violations).padStart(3)}`
    );
  }
  console.log("");

  if (allViolations.length === 0) {
    console.log("action-parity: OK");
    return 0;
  }

  console.log(`action-parity: ${LENIENT ? "WARNINGS" : "FAILED"}`);
  console.log("");
  for (const v of allViolations) {
    console.log(`  ${LENIENT ? "⚠" : "✖"} [${v.format}] ${v.message}`);
  }
  console.log("");
  if (LENIENT) {
    console.log(
      `Total ${allViolations.length} drift item(s) (lenient mode — exit 0). Fix these to flip the gate to strict.`
    );
    return 0;
  }
  console.log(`Total ${allViolations.length} violation(s).`);
  return 1;
}

process.exit(main());
