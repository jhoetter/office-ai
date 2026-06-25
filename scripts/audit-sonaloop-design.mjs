#!/usr/bin/env node
import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const ROOT = resolve(fileURLToPath(import.meta.url), "..", "..");
const DESIGN_ROOT = resolve(ROOT, "..", "sonaloop-design");
const OUT_PATH = join(ROOT, "docs", "sonaloop-design-adoption.md");

const SOURCE_ROOTS = ["apps/web/app", "packages/ui/src", "packages/react-editors/src"];
const SOURCE_EXTS = new Set([".ts", ".tsx", ".js", ".jsx"]);
const SKIP_DIRS = new Set(["node_modules", "dist", ".next", "coverage", ".turbo"]);
const ICON_IMPORT_SOURCES = new Set(["lucide-react", "@officeai/ui/sonaloop-icons", "../sonaloop-icons"]);

const ICON_ALIASES = {
  AlertTriangle: "alert",
  Check: "check",
  CheckCircle2: "check",
  ChevronDown: "chevron",
  ChevronLeft: "back",
  ChevronRight: "chevronRight",
  ChevronUp: "chevronUp",
  Clock3: "clock",
  Command: "command",
  ExternalLink: "external",
  FileArchive: "archive",
  FileCode: "codeFile",
  FileImage: "imageFile",
  FileSpreadsheet: "spreadsheetFile",
  FileText: "documentFile",
  FileType2: "pdfFile",
  FolderOpen: "folderOpen",
  Info: "info",
  LayoutTemplate: "layoutTemplate",
  Link: "link",
  Loader2: "loader",
  Maximize2: "expand",
  Minimize2: "collapse",
  Monitor: "monitor",
  Moon: "moon",
  Plus: "plus",
  Presentation: "presentationFile",
  Search: "search",
  Sun: "sun",
  Trash2: "trash",
  Upload: "upload",
  X: "close",
  XCircle: "close",
};

function walk(dir, out = []) {
  let entries;
  try {
    entries = readdirSync(dir);
  } catch {
    return out;
  }
  for (const entry of entries) {
    if (SKIP_DIRS.has(entry)) continue;
    const path = join(dir, entry);
    const st = statSync(path);
    if (st.isDirectory()) walk(path, out);
    else if (st.isFile() && SOURCE_EXTS.has(path.slice(path.lastIndexOf(".")))) out.push(path);
  }
  return out;
}

function extractLucideImports(file) {
  const text = readFileSync(file, "utf8");
  const imports = [];
  const blocks = [];
  let current = [];
  for (const line of text.split(/\r?\n/)) {
    if (/^\s*import\s/.test(line)) {
      if (current.length > 0) current = [];
      current.push(line);
    } else if (current.length > 0) {
      current.push(line);
    }

    if (current.length === 0) continue;
    const block = current.join("\n");
    const source = block.match(/from\s*["']([^"']+)["']/)?.[1];
    if (source && ICON_IMPORT_SOURCES.has(source)) {
      blocks.push(block);
      current = [];
    } else if (source) {
      current = [];
    }
  }

  for (const block of blocks) {
    const match = /import\s*\{([\s\S]*?)\}\s*from\s*["']([^"']+)["']/.exec(block);
    if (!match || !ICON_IMPORT_SOURCES.has(match[2])) continue;
    for (const raw of match[1].split(",")) {
      const item = raw.trim();
      if (!item || item.startsWith("type ")) continue;
      const name = item.split(/\s+as\s+/)[0]?.trim();
      if (name) imports.push(name);
    }
  }
  return imports;
}

async function loadDesignIcons() {
  const dataPath = join(DESIGN_ROOT, "icons.data.mjs");
  if (!existsSync(dataPath)) return { regular: {}, hifi: {} };
  const mod = await import(pathToFileURL(dataPath).href);
  return { regular: mod.regular ?? {}, hifi: mod.hifi ?? {} };
}

function normalize(value) {
  return String(value)
    .replace(/[^a-z0-9]/gi, "")
    .toLowerCase();
}

function categoryFor(name) {
  if (/file|archive|folder|upload|download|export|image|presentation|spreadsheet|pdf/i.test(name)) {
    return "file-format/export";
  }
  if (/bold|italic|underline|strikethrough|palette|highlighter|type|align|list/i.test(name)) {
    return "text-formatting";
  }
  if (/table|grid|rows|columns|merge|filter|sort|function|chart/i.test(name)) return "tables-grid";
  if (/slide|shape|layout|connector|maximize|minimize|monitor/i.test(name)) return "slides-shapes";
  if (/pdf|book|bookmark|message|comment|annotation|link/i.test(name)) return "pdf-review";
  if (/search|command|loader|clock|refresh|settings|x|check|alert|info|trash|plus/i.test(name)) {
    return "chrome-command";
  }
  return "general";
}

function findIconKey(lucideName, regularKeys, hifiKeys) {
  const alias = ICON_ALIASES[lucideName];
  if (alias && regularKeys.includes(alias)) return { key: alias, status: "mapped-existing" };
  if (alias && hifiKeys.includes(alias)) return { key: alias, status: "mapped-existing-hifi" };
  const normalized = normalize(lucideName);
  const exactRegular = regularKeys.find((key) => normalize(key) === normalized);
  if (exactRegular) return { key: exactRegular, status: "mapped-existing" };
  const exactHifi = hifiKeys.find((key) => normalize(key) === normalized);
  if (exactHifi) return { key: exactHifi, status: "mapped-existing-hifi" };
  const proposed = alias ?? lucideName.charAt(0).toLowerCase() + lucideName.slice(1);
  return { key: proposed, status: "gap" };
}

function rel(path) {
  return relative(ROOT, path).replace(/\\/g, "/");
}

function mdEscape(value) {
  return String(value).replace(/\|/g, "\\|");
}

function formatGeneratedMarkdown() {
  const prettierBin = join(
    ROOT,
    "node_modules",
    ".bin",
    process.platform === "win32" ? "prettier.cmd" : "prettier"
  );
  const command = existsSync(prettierBin) ? prettierBin : "prettier";
  const result = spawnSync(command, ["--write", rel(OUT_PATH)], {
    cwd: ROOT,
    encoding: "utf8",
  });
  if (result.status !== 0) {
    const message = result.stderr || result.stdout || "unknown error";
    console.warn(`warning: could not format ${rel(OUT_PATH)}: ${message.trim()}`);
  }
}

async function main() {
  const files = SOURCE_ROOTS.flatMap((root) => walk(join(ROOT, root))).sort();
  const usage = new Map();
  for (const file of files) {
    for (const name of extractLucideImports(file)) {
      const current = usage.get(name) ?? { count: 0, files: new Set() };
      current.count += 1;
      current.files.add(rel(file));
      usage.set(name, current);
    }
  }

  const { regular, hifi } = await loadDesignIcons();
  const regularKeys = Object.keys(regular).sort();
  const hifiKeys = Object.keys(hifi).sort();
  const rows = [...usage.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([name, data]) => {
      const mapping = findIconKey(name, regularKeys, hifiKeys);
      return {
        lucide: name,
        count: data.count,
        files: [...data.files].sort(),
        category: categoryFor(name),
        key: mapping.key,
        status: mapping.status,
      };
    });
  const gaps = rows.filter((row) => row.status === "gap");
  const mapped = rows.length - gaps.length;

  const lines = [];
  lines.push("# Sonaloop design adoption audit");
  lines.push("");
  lines.push("Generated by `pnpm design:audit`.");
  lines.push("");
  lines.push("## Summary");
  lines.push("");
  lines.push(`- Office-AI icon imports scanned: ${rows.length} unique icon names.`);
  lines.push(`- Already mapped to ` + "`sonaloop-design`" + `: ${mapped}.`);
  lines.push(`- Missing from ` + "`sonaloop-design`" + `: ${gaps.length}.`);
  lines.push(`- Source roots: ${SOURCE_ROOTS.map((root) => `\`${root}\``).join(", ")}.`);
  lines.push("");
  lines.push("## Command palette API diff");
  lines.push("");
  lines.push("| Concern | office-ai today | sonaloop-design target | Migration note |");
  lines.push("| --- | --- | --- | --- |");
  lines.push(
    "| Data model | `PaletteCommand` from action catalogue | `CommandGroup` / `CommandItem` | Add adapter from catalogue actions to grouped command items. |"
  );
  lines.push(
    "| Icons | `@officeai/ui/sonaloop-icons` adapter components | `IconKey` strings rendered by `sonaloop-design` | Keep call-site names stable while the adapter maps to shared icons. |"
  );
  lines.push(
    "| Hotkey | Local shell handler | Built-in CMD+K/Ctrl-K in `CommandPalette` | Let the shared palette own the global hotkey. |"
  );
  lines.push(
    "| Navigation | Local Next routing/action callbacks | `SonaloopLinkProvider` link adapter plus `onSelect` | Provide a Next router-backed link component once at shell root. |"
  );
  lines.push(
    "| Search | Local filtering/recent behavior | Client filter plus optional `onSearch` | Preserve recents as an extra generated group if still required. |"
  );
  lines.push("");
  lines.push("## Token mapping");
  lines.push("");
  lines.push("| office-ai source | sonaloop-design target | Recommendation |");
  lines.push("| --- | --- | --- |");
  lines.push(
    "| `@officeai/design-tokens/src/colors.ts` | `sonaloop-design/styles/tokens.css` `--sl-*` aliases | Keep `@officeai/design-tokens` only as a temporary compatibility re-export/mapping layer. |"
  );
  lines.push(
    "| `@officeai/design-tokens/src/tailwind-preset.ts` | `sonaloop-design` app CSS preset | Replace direct token values with shared CSS imports and semantic `.sl-*` classes. |"
  );
  lines.push(
    "| `apps/web/app/globals.css` local variables | `--sl-bg`, `--sl-surface`, `--sl-ink`, `--sl-muted`, `--sl-line`, `--sl-accent` | Alias old names to `--sl-*` during migration, then remove old names. |"
  );
  lines.push(
    "| `packages/ui` Tailwind utility strings | `styles/components.css` component classes | Migrate common controls first: buttons, inputs, badges, empty states, command palette. |"
  );
  lines.push("");
  lines.push("## Icon mapping table");
  lines.push("");
  lines.push("| Lucide import | Uses | Category | sonaloop IconKey | Status | Files |");
  lines.push("| --- | ---: | --- | --- | --- | --- |");
  for (const row of rows) {
    lines.push(
      `| \`${row.lucide}\` | ${row.count} | ${row.category} | \`${row.key}\` | ${row.status} | ${mdEscape(
        row.files.slice(0, 6).join("<br>")
      )}${row.files.length > 6 ? "<br>..." : ""} |`
    );
  }
  lines.push("");
  lines.push("## Gap list");
  lines.push("");
  if (gaps.length === 0) {
    lines.push("No icon gaps remain.");
  } else {
    for (const row of gaps) {
      lines.push(`- \`${row.lucide}\` -> proposed \`${row.key}\` (${row.category}).`);
    }
  }
  lines.push("");
  lines.push("## Migration recommendation");
  lines.push("");
  lines.push(
    "1. Add the missing Office-domain icons to `sonaloop-design/icons.data.mjs` under the taxonomy in `docs/sonaloop-office-icon-taxonomy.md`."
  );
  lines.push("2. Generate and publish/consume the `sonaloop-design` React icon layer.");
  lines.push(
    "3. Keep Office-AI call sites on `@officeai/ui/sonaloop-icons`; add new glyph geometry only in `sonaloop-design/icons.data.mjs`."
  );
  lines.push(
    "4. Keep `@officeai/design-tokens` as a compatibility mapping only until app-shell and primitive classes consume `sonaloop-design` directly."
  );
  lines.push("");

  mkdirSync(dirname(OUT_PATH), { recursive: true });
  writeFileSync(OUT_PATH, `${lines.join("\n")}\n`);
  formatGeneratedMarkdown();
  console.log(`wrote ${rel(OUT_PATH)} (${rows.length} icons, ${gaps.length} gaps)`);
}

main().catch((err) => {
  console.error(err instanceof Error ? err.message : String(err));
  process.exit(1);
});
