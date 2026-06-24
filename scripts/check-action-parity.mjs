#!/usr/bin/env node
/**
 * Action parity gate.
 *
 * For each format `F` in {docx, xlsx, pptx, pdf}, verifies that the
 * action catalogue at `packages/F/src/actions/catalogue.ts` covers
 * every command-bus handler registered in
 * `packages/F/src/commands/`. Catches the most expensive class of
 * drift in the codebase: a new bus command lands without a CLI
 * subcommand, MCP tool, or Cmd+K palette entry, so it's reachable by
 * one surface but invisible to humans or agents elsewhere.
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
    /**
     * Editor source roots scanned for ribbon/palette wiring of
     * `toolbar`-surfaced catalogue entries. The toolbar parity check
     * looks for `type: "<format>:..."` literals inside `applyCommand`
     * callsites here as a proxy for "this action is wired into the
     * UI". Direct opens of dialogs that ultimately dispatch the
     * command count too — the literal still appears once in the
     * editor source.
     */
    uiDirs: [
      "apps/web/app/editor",
      // Shared helpers that dispatch DOCX commands from outside the
      // ribbon code path (image insertion via drop/paste, page-break
      // keymap, etc.). The toolbar coverage check only cares that
      // *some* UI surface dispatches the command; these helpers are
      // mounted by the editor on every render.
      "apps/web/app/lib/image-insert.ts",
      "apps/web/app/lib/page-keymap.ts",
    ],
  },
  {
    name: "xlsx",
    handlerDirs: ["packages/xlsx/src/commands"],
    cataloguePath: "packages/xlsx/src/actions/catalogue.ts",
    uiDirs: ["apps/web/app/xlsx-editor"],
  },
  {
    name: "pptx",
    handlerDirs: ["packages/pptx/src/commands"],
    cataloguePath: "packages/pptx/src/actions/catalogue.ts",
    uiDirs: ["apps/web/app/pptx-editor"],
  },
  {
    name: "pdf",
    handlerDirs: ["packages/pdf/src/commands"],
    cataloguePath: "packages/pdf/src/actions/catalogue.ts",
    // The PDF reader lives at `apps/web/app/pdf-viewer/`. The historical
    // path `apps/web/app/pdf-editor/` was a phantom — the script reported
    // `pdf ui-dispatched=0` because no such directory exists. Phase 9d
    // (PDF parity sweep) replaces it with the real reader root so the
    // gate actually sees the wired PDF buttons.
    uiDirs: ["apps/web/app/pdf-viewer"],
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
 * Pull every catalogue entry's structural metadata from a single
 * catalogue.ts file. Uses a deliberately narrow regex over the
 * canonical declaration shape; if someone reformats the file past
 * what the regex matches, the parity check will yell with a clear
 * "missing action metadata" line so the regex can be tightened.
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
    const declaredFormat = parseStringField(block, "format");
    const surfaces = parseSurfaces(block);
    const agentCallable = parseBooleanField(block, "agentCallable");
    const webCallable = parseBooleanField(block, "webCallable");
    const cliCallable = parseBooleanField(block, "cliCallable");
    const requiresReview = parseBooleanField(block, "requiresReview");
    const supportsDryRun = parseBooleanField(block, "supportsDryRun");
    const supportsDiff = parseBooleanField(block, "supportsDiff");
    const commandSchema = parseStringField(block, "commandSchema");
    const hidden = /\bhidden\s*:\s*\{/.test(block);
    const hasArgs = /\bargs\s*:/.test(block);
    const hasBuildPayload = /\bbuildPayload\s*:/.test(block);
    entries.push({
      id,
      commandType,
      declaredFormat,
      surfaces,
      agentCallable,
      webCallable,
      cliCallable,
      requiresReview,
      supportsDryRun,
      supportsDiff,
      commandSchema,
      hidden,
      hasArgs,
      hasBuildPayload,
    });
  }
  return entries;
}

function parseCommandType(block) {
  const m = block.match(/\bcommandType:\s*(?:"([^"]+)"|null)/);
  if (!m) return undefined;
  return m[1] ?? null;
}

function parseSurfaces(block) {
  // surfaces: ["palette", "toolbar"]
  const m = block.match(/\bsurfaces:\s*\[([^\]]*)\]/);
  if (!m) return [];
  return [...m[1].matchAll(/"([a-zA-Z]+)"/g)].map((x) => x[1]);
}

function parseBooleanField(block, field) {
  const m = block.match(new RegExp(`\\b${field}:\\s*(true|false)`));
  if (!m) return undefined;
  return m[1] === "true";
}

function parseStringField(block, field) {
  const m = block.match(new RegExp(`\\b${field}:\\s*"([^"]+)"`));
  return m ? m[1] : undefined;
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

/**
 * Scan `dirs` recursively for any UI source file (`.tsx`, `.ts`)
 * that mentions a `type: "<format>:..."` literal — i.e. dispatches
 * the command via `applyCommand`. Returns the set of distinct
 * `<format>:...` types found.
 */
function extractUiDispatchedTypes(format, uiDirs) {
  const types = new Set();
  // Two patterns count: the canonical `type: "<format>:..."`
  // discriminator (used by `applyCommand({ type: ..., payload: ... })`),
  // and the loose first-arg form `"<format>:..."` that helpers like
  // `dispatchAny(...)` and palette runners pass through. Together
  // they cover every UI dispatch path we currently use.
  const reType = new RegExp(`\\btype:\\s*"(${format}:[a-z][a-z0-9-]*)"`, "g");
  const reLoose = new RegExp(`"(${format}:[a-z][a-z0-9-]*)"`, "g");
  const visit = (file, contents) => {
    if (file.endsWith(".test.ts") || file.endsWith(".test.tsx")) return;
    let m;
    while ((m = reType.exec(contents)) !== null) {
      types.add(m[1]);
    }
    while ((m = reLoose.exec(contents)) !== null) {
      types.add(m[1]);
    }
  };
  for (const dir of uiDirs) {
    const abs = join(ROOT, dir);
    let st;
    try {
      st = statSync(abs);
    } catch {
      continue;
    }
    if (st.isFile()) {
      visit(abs, readFileSync(abs, "utf8"));
    } else if (st.isDirectory()) {
      walkUiFiles(abs, visit);
    }
  }
  return types;
}

function walkUiFiles(dir, visit) {
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
      walkUiFiles(full, visit);
    } else if (st.isFile() && (full.endsWith(".ts") || full.endsWith(".tsx"))) {
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
 *  4. Every catalogue entry must make explicit agent/web/CLI,
 *     review, dry-run, diff, format, and command-schema decisions.
 *  5. UI surfaces stay UI-only: `surfaces` may not contain `cli`.
 *  6. Generated MCP/CLI bindings require catalogue-owned args plus
 *     buildPayload; terminal-only conveniences may remain
 *     `cliCallable` without becoming `agentCallable`.
 */
function checkFormat(format, handlerTypes, entries, i18nKeys, uiDispatchedTypes) {
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
    violations.push(...checkActionMetadata(format, e));
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

  function pushMetadataViolation(e, message) {
    return {
      format,
      kind: "invalid-action-metadata",
      message: `catalogue action "${e.id}" has invalid metadata — ${message}`,
    };
  }

  function checkActionMetadata(formatName, e) {
    const out = [];
    const booleanFields = [
      "agentCallable",
      "webCallable",
      "cliCallable",
      "requiresReview",
      "supportsDryRun",
      "supportsDiff",
    ];
    const validUiSurfaces = new Set(["palette", "toolbar", "contextMenu"]);
    const validSchemas = new Set(["catalogue-args", "custom", "none"]);
    const isBusBacked = e.commandType !== null && e.commandType !== undefined;

    if (e.declaredFormat !== formatName) {
      out.push(pushMetadataViolation(e, `format must be "${formatName}"`));
    }
    for (const field of booleanFields) {
      if (typeof e[field] !== "boolean") {
        out.push(pushMetadataViolation(e, `${field} must be explicitly true or false`));
      }
    }
    if (!validSchemas.has(e.commandSchema)) {
      out.push(pushMetadataViolation(e, `commandSchema must be one of ${[...validSchemas].join(", ")}`));
    }
    for (const surface of e.surfaces) {
      if (surface === "cli") {
        out.push(pushMetadataViolation(e, 'surfaces is UI-only; use cliCallable instead of "cli"'));
      } else if (!validUiSurfaces.has(surface)) {
        out.push(pushMetadataViolation(e, `unknown UI surface "${surface}"`));
      }
    }

    if (isBusBacked) {
      if (e.requiresReview !== true) {
        out.push(pushMetadataViolation(e, "bus-backed mutations must set requiresReview: true"));
      }
      if (e.supportsDiff !== true) {
        out.push(pushMetadataViolation(e, "bus-backed mutations must set supportsDiff: true"));
      }
      if (e.commandSchema === "none") {
        out.push(pushMetadataViolation(e, 'bus-backed actions cannot use commandSchema: "none"'));
      }
    } else {
      if (e.requiresReview === true) {
        out.push(pushMetadataViolation(e, "non-bus actions must set requiresReview: false"));
      }
      if (e.supportsDiff === true) {
        out.push(pushMetadataViolation(e, "non-bus actions must set supportsDiff: false"));
      }
      if (e.commandSchema !== "none") {
        out.push(pushMetadataViolation(e, 'non-bus actions must use commandSchema: "none"'));
      }
    }

    if (e.commandSchema === "catalogue-args" && (!e.hasArgs || !e.hasBuildPayload)) {
      out.push(
        pushMetadataViolation(e, 'commandSchema: "catalogue-args" requires both args and buildPayload')
      );
    }
    if (e.agentCallable) {
      if (e.hidden) {
        out.push(pushMetadataViolation(e, "hidden actions cannot be agentCallable"));
      }
      if (!isBusBacked) {
        out.push(pushMetadataViolation(e, "agentCallable requires a bus commandType"));
      }
      if (e.commandSchema !== "catalogue-args") {
        out.push(pushMetadataViolation(e, 'agentCallable requires commandSchema: "catalogue-args"'));
      }
    }
    if (e.webCallable && e.hidden) {
      out.push(pushMetadataViolation(e, "hidden actions cannot be webCallable"));
    }
    if (e.cliCallable && e.hidden) {
      out.push(pushMetadataViolation(e, "hidden actions cannot be cliCallable"));
    }
    if (!e.hidden && e.surfaces.length > 0 && e.webCallable !== true) {
      out.push(pushMetadataViolation(e, "actions with UI surfaces must set webCallable: true"));
    }
    if (e.webCallable && e.surfaces.length === 0) {
      out.push(pushMetadataViolation(e, "webCallable actions must declare at least one UI surface"));
    }

    return out;
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
    const localId = e.id.startsWith(`${format}.`) ? e.id.slice(format.length + 1) : e.id;
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

  // Toolbar coverage: every catalogue entry whose `surfaces`
  // include "toolbar" and that has a `commandType` should appear
  // somewhere in the matching format's editor source — otherwise
  // the catalogue claims a ribbon affordance that the UI does not
  // actually expose. We use `applyCommand({ type: "<format>:..." })`
  // literals as the proxy: every dispatched command in the editor
  // (whether from a button, a splitter menu item, a dialog submit,
  // or a palette runner) carries this literal in source.
  //
  // Allowlisted exceptions live in `TOOLBAR_COVERAGE_ALLOWLIST`
  // below for entries that intentionally surface only via menus
  // built up indirectly (e.g. a splitter that fans out to a
  // sub-action whose dispatch happens in a shared helper). New
  // catalogue entries should not be added here without a comment
  // explaining why the heuristic doesn't apply.
  for (const e of entries) {
    if (e.hidden) continue;
    if (!e.surfaces.includes("toolbar")) continue;
    if (!e.commandType) continue;
    if (TOOLBAR_COVERAGE_ALLOWLIST.has(e.id)) continue;
    if (!uiDispatchedTypes.has(e.commandType)) {
      violations.push({
        format,
        kind: "missing-toolbar-wiring",
        message: `catalogue action "${e.id}" declares surfaces:["toolbar"] but no UI source under apps/web/app/${format}-editor/ (or apps/web/app/editor for docx) dispatches "${e.commandType}" — add a ribbon button/menu, or remove "toolbar" from surfaces, or allowlist in scripts/check-action-parity.mjs`,
      });
    }
  }

  return violations;
}

/**
 * Catalogue ids whose `toolbar` surface is wired indirectly (via a
 * shared helper that doesn't carry the `type: "..."` literal in the
 * editor source). Keep this list short and well-commented; if it
 * starts growing, refactor the helper to dispatch directly so the
 * coverage check stays honest.
 */
const TOOLBAR_COVERAGE_ALLOWLIST = new Set([
  // Text-formatting commands flow through `TextFormatProvider` /
  // `format-range` from `@officeai/text-formatting`; the literal
  // lives in the package, not in the editor source. The format-bar
  // is the only consumer and ships with every editor.
  "docx.format-range",
  "xlsx.format-range",
  "pptx.format-range",
]);

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
    const uiDispatchedTypes = extractUiDispatchedTypes(f.name, f.uiDirs);
    const violations = checkFormat(f.name, handlerTypes, entries, i18nKeys, uiDispatchedTypes);
    summary.push({
      format: f.name,
      handlers: handlerTypes.size,
      catalogue: entries.length,
      agentCallable: entries.filter((e) => e.agentCallable === true).length,
      webCallable: entries.filter((e) => e.webCallable === true).length,
      cliCallable: entries.filter((e) => e.cliCallable === true).length,
      uiDispatched: uiDispatchedTypes.size,
      violations: violations.length,
    });
    allViolations.push(...violations);
  }

  console.log("action-parity check");
  console.log("───────────────────");
  for (const s of summary) {
    console.log(
      `  ${s.format.padEnd(6)} handlers=${String(s.handlers).padStart(3)}  catalogue=${String(s.catalogue).padStart(3)}  agent=${String(s.agentCallable).padStart(3)}  web=${String(s.webCallable).padStart(3)}  cli=${String(s.cliCallable).padStart(3)}  ui-dispatched=${String(s.uiDispatched).padStart(3)}  violations=${String(s.violations).padStart(3)}`
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
