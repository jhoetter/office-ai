#!/usr/bin/env node
/**
 * Generate the cross-surface parity scorecard for MCP, Web and CLI.
 *
 * The scorecard is intentionally source-derived: action catalogues are
 * the product-facing capability list; command handlers prove bus
 * support; MCP registrations prove hand-rolled tools; UI source
 * literals prove editor dispatch wiring.
 */

import { existsSync, mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(fileURLToPath(import.meta.url), "..", "..");
const CHECK = process.argv.includes("--check");
const MD_OUT = argValue("--md-out") ?? "docs/cross-surface-parity-scorecard.md";
const JSON_OUT = argValue("--json-out") ?? "docs/cross-surface-parity-scorecard.json";

const FORMATS = [
  {
    name: "docx",
    handlerDirs: ["packages/docx/src/commands"],
    cataloguePath: "packages/docx/src/actions/catalogue.ts",
    uiDirs: ["apps/web/app/editor", "apps/web/app/lib/image-insert.ts", "apps/web/app/lib/page-keymap.ts"],
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
    uiDirs: ["apps/web/app/pdf-viewer"],
  },
];

const MCP_TOOL_ALIASES = {
  "docx.apply": ["apply_command"],
  "docx.create": ["create_document"],
  "docx.read": ["docx_get_text", "get_document_projection"],
  "xlsx.apply": ["apply_command"],
  "xlsx.create": ["create_document"],
  "xlsx.read": ["xlsx_get_text", "get_document_projection"],
  "pptx.apply": ["apply_command"],
  "pptx.create": ["create_document"],
  "pptx.read": ["pptx_get_text", "get_document_projection"],
  "pdf.create": ["create_document"],
};

function argValue(name) {
  const idx = process.argv.indexOf(name);
  return idx === -1 ? undefined : process.argv[idx + 1];
}

function walkFiles(start, predicate, visit) {
  let st;
  try {
    st = statSync(start);
  } catch {
    return;
  }
  if (st.isFile()) {
    if (predicate(start)) visit(start, readFileSync(start, "utf8"));
    return;
  }
  if (!st.isDirectory()) return;
  for (const entry of readdirSync(start)) {
    walkFiles(join(start, entry), predicate, visit);
  }
}

function extractHandlerTypes(format, handlerDirs) {
  const types = new Set();
  for (const dir of handlerDirs) {
    walkFiles(
      join(ROOT, dir),
      (file) => file.endsWith(".ts") && !file.endsWith(".test.ts"),
      (_file, contents) => {
        const re = new RegExp(`\\btype:\\s*"(${format}:[a-z][a-z0-9-]*)"`, "g");
        let m;
        while ((m = re.exec(contents)) !== null) types.add(m[1]);
      }
    );
  }
  return types;
}

function extractUiDispatchedTypes(format, uiDirs) {
  const types = new Set();
  const reType = new RegExp(`\\btype:\\s*"(${format}:[a-z][a-z0-9-]*)"`, "g");
  const reLoose = new RegExp(`"(${format}:[a-z][a-z0-9-]*)"`, "g");
  const visit = (_file, contents) => {
    let m;
    while ((m = reType.exec(contents)) !== null) types.add(m[1]);
    while ((m = reLoose.exec(contents)) !== null) types.add(m[1]);
  };
  for (const target of uiDirs) {
    walkFiles(
      join(ROOT, target),
      (file) => (file.endsWith(".ts") || file.endsWith(".tsx")) && !file.includes(".test."),
      visit
    );
  }
  return types;
}

function extractMcpToolNames() {
  const src = readFileSync(join(ROOT, "packages/agent/src/mcp.ts"), "utf8");
  const names = new Set();
  const re = /server\.registerTool\(\s*"([^"]+)"/g;
  let m;
  while ((m = re.exec(src)) !== null) names.add(m[1]);
  return names;
}

function extractCatalogueEntries(format, cataloguePath) {
  const abs = join(ROOT, cataloguePath);
  const src = readFileSync(abs, "utf8");
  const entries = [];
  const idRe = /\{\s*id:\s*"([^"]+)"/g;
  const positions = [];
  let m;
  while ((m = idRe.exec(src)) !== null) {
    positions.push({ id: m[1], start: m.index });
  }
  for (let i = 0; i < positions.length; i++) {
    const start = positions[i].start;
    const end = i + 1 < positions.length ? positions[i + 1].start : src.length;
    const block = src.slice(start, end);
    entries.push({
      id: positions[i].id,
      format,
      operation: operationKind(positions[i].id, parseStringField(block, "section")),
      section: parseStringField(block, "section") ?? "Unsectioned",
      label: parseStringField(block, "label") ?? positions[i].id,
      description: parseStringField(block, "description") ?? "",
      commandType: parseCommandType(block),
      surfaces: parseSurfaces(block),
      agentCallable: parseBooleanField(block, "agentCallable") === true,
      webCallable: parseBooleanField(block, "webCallable") === true,
      cliCallable: parseBooleanField(block, "cliCallable") === true,
      requiresReview: parseBooleanField(block, "requiresReview") === true,
      supportsDryRun: parseBooleanField(block, "supportsDryRun") === true,
      supportsDiff: parseBooleanField(block, "supportsDiff") === true,
      commandSchema: parseStringField(block, "commandSchema") ?? "unknown",
      hidden: /\bhidden\s*:\s*\{/.test(block),
      hasArgs: /\bargs\s*:/.test(block),
      hasBuildPayload: /\bbuildPayload\s*:/.test(block),
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
  const m = block.match(new RegExp(`\\b${field}:\\s*"([^"]*)"`));
  return m ? m[1] : undefined;
}

function operationKind(id, section = "") {
  const tokens = new Set(
    `${id} ${section}`
      .toLowerCase()
      .split(/[^a-z0-9]+/)
      .filter(Boolean)
  );
  if (hasAny(tokens, ["export", "save", "download", "markdown"])) return "export";
  if (hasAny(tokens, ["search", "find"])) return "search";
  if (hasAny(tokens, ["annotation", "annotations", "comment", "comments", "note", "highlight", "redact"])) {
    return "annotate";
  }
  if (hasAny(tokens, ["read", "inspect", "metadata", "outline", "list", "get", "extract"])) return "read";
  return "edit";
}

function hasAny(tokens, values) {
  return values.some((value) => tokens.has(value));
}

function mcpToolName(actionId) {
  const dot = actionId.indexOf(".");
  if (dot === -1) return actionId.replace(/[-.]/g, "_");
  const format = actionId.slice(0, dot);
  const local = actionId.slice(dot + 1).replace(/[-.]/g, "_");
  return `${format}_${local}`;
}

function cliCommand(actionId) {
  const dot = actionId.indexOf(".");
  if (dot === -1) return `office-agent ${actionId}`;
  return `office-agent ${actionId.slice(0, dot)} ${actionId.slice(dot + 1)}`;
}

function surfaceStatus(surface, entry, context) {
  if (entry.hidden) {
    return {
      status: "not_relevant",
      evidence: ["catalogue:hidden"],
      note: "Hidden catalogue entry; wrapper or internal affordance owns the public surface.",
    };
  }

  const hasHandler = entry.commandType ? context.handlerTypes.has(entry.commandType) : false;
  const toolName = mcpToolName(entry.id);
  const mcpTools = matchingMcpTools(entry, context.mcpToolNames);
  const hasMcpTool = mcpTools.length > 0;
  const uiDispatch = entry.commandType ? context.uiDispatchedTypes.has(entry.commandType) : false;

  if (surface === "mcp") {
    if (entry.commandType && hasHandler) {
      return {
        status: "complete",
        evidence: [`plan_command/action_id:${entry.id}`, `handler:${entry.commandType}`],
        note: "Canonical MCP command lifecycle can plan, preview and apply this bus-backed action.",
      };
    }
    if (hasMcpTool || entry.agentCallable) {
      return {
        status: hasMcpTool || isGeneratedMcpBindable(entry) ? "complete" : "partial",
        evidence: hasMcpTool ? mcpTools.map((name) => `tool:${name}`) : ["catalogue:agentCallable"],
        note: hasMcpTool
          ? "Hand-rolled MCP tool is registered."
          : "Marked agent-callable but not proven as a registered MCP tool by static scan.",
      };
    }
    return {
      status: "missing",
      evidence: ["no mcp tool or command lifecycle coverage found"],
      note: "Public catalogue action has no MCP surface yet.",
    };
  }

  if (surface === "web") {
    if (entry.webCallable && (entry.surfaces.length > 0 || uiDispatch || !entry.commandType)) {
      return {
        status: uiDispatch || !entry.commandType ? "complete" : "partial",
        evidence: [
          `catalogue:webCallable`,
          entry.surfaces.length ? `surfaces:${entry.surfaces.join(",")}` : "surfaces:none",
          ...(uiDispatch ? [`ui-dispatch:${entry.commandType}`] : []),
        ],
        note:
          uiDispatch || !entry.commandType
            ? "Web catalogue and dispatch evidence are present."
            : "Catalogue declares a web surface; dispatch literal was not found by the static scanner.",
      };
    }
    if (uiDispatch) {
      return {
        status: "partial",
        evidence: [`ui-dispatch:${entry.commandType}`],
        note: "Editor source dispatches this command, but the catalogue does not mark it webCallable.",
      };
    }
    return {
      status: "missing",
      evidence: ["catalogue:webCallable=false"],
      note: "No web editor surface is declared.",
    };
  }

  if (surface === "cli") {
    if (entry.cliCallable) {
      return {
        status: "complete",
        evidence: [`catalogue:cliCallable`, cliCommand(entry.id)],
        note: "Catalogue marks this action as available to the terminal CLI.",
      };
    }
    return {
      status: "missing",
      evidence: ["catalogue:cliCallable=false"],
      note: "No CLI surface is declared.",
    };
  }

  throw new Error(`Unknown surface ${surface}`);
}

function matchingMcpTools(entry, mcpToolNames) {
  const candidates = [mcpToolName(entry.id), ...(MCP_TOOL_ALIASES[entry.id] ?? [])];
  return candidates.filter((name, idx) => candidates.indexOf(name) === idx && mcpToolNames.has(name));
}

function isGeneratedMcpBindable(entry) {
  return (
    entry.agentCallable &&
    !entry.hidden &&
    entry.commandType !== null &&
    entry.commandType !== undefined &&
    entry.commandSchema === "catalogue-args" &&
    entry.hasArgs &&
    entry.hasBuildPayload
  );
}

function buildScorecard() {
  const mcpToolNames = extractMcpToolNames();
  const formats = [];
  const entries = [];
  for (const format of FORMATS) {
    const handlerTypes = extractHandlerTypes(format.name, format.handlerDirs);
    const uiDispatchedTypes = extractUiDispatchedTypes(format.name, format.uiDirs);
    const catalogue = extractCatalogueEntries(format.name, format.cataloguePath);
    const context = { handlerTypes, uiDispatchedTypes, mcpToolNames };
    const scored = catalogue.map((entry) => ({
      ...entry,
      surfacesScore: {
        mcp: surfaceStatus("mcp", entry, context),
        web: surfaceStatus("web", entry, context),
        cli: surfaceStatus("cli", entry, context),
      },
    }));
    formats.push({
      format: format.name,
      handlers: handlerTypes.size,
      actions: catalogue.length,
      uiDispatched: uiDispatchedTypes.size,
      mcpToolsSeen: [...mcpToolNames].filter((name) => name.startsWith(`${format.name}_`)).length,
      summary: summarize(scored),
    });
    entries.push(...scored);
  }
  return {
    schema: "office-ai/cross-surface-parity-scorecard@1",
    sources: {
      actionCatalogues: FORMATS.map((f) => f.cataloguePath),
      handlerDirs: FORMATS.flatMap((f) => f.handlerDirs),
      mcpTools: "packages/agent/src/mcp.ts",
      webDispatchRoots: Object.fromEntries(FORMATS.map((f) => [f.name, f.uiDirs])),
    },
    statusVocabulary: ["complete", "partial", "missing", "not_relevant"],
    formats,
    entries,
  };
}

function summarize(entries) {
  const out = {};
  for (const surface of ["mcp", "web", "cli"]) {
    out[surface] = { complete: 0, partial: 0, missing: 0, not_relevant: 0 };
    for (const entry of entries) {
      out[surface][entry.surfacesScore[surface].status] += 1;
    }
  }
  return out;
}

function renderMarkdown(scorecard) {
  const lines = [];
  lines.push("# Cross-Surface Parity Scorecard");
  lines.push("");
  lines.push(
    "Generated by `pnpm scorecard` from action catalogues, command handlers, MCP tool registrations and web editor dispatch literals."
  );
  lines.push(
    "The check mode (`pnpm scorecard:check`) verifies that this file and the JSON companion stay reproducible."
  );
  lines.push("");
  lines.push("## Status Vocabulary");
  lines.push("");
  lines.push("| Status | Meaning |");
  lines.push("| --- | --- |");
  lines.push("| `complete` | Source evidence says the capability is available on that surface. |");
  lines.push("| `partial` | A declaration exists, but static evidence is incomplete or indirect. |");
  lines.push(
    "| `missing` | The capability is public in another surface but has no declared coverage here. |"
  );
  lines.push(
    "| `not_relevant` | Hidden/internal catalogue entry or deliberately wrapped by another action. |"
  );
  lines.push("");
  lines.push("## Release Gate");
  lines.push("");
  lines.push("- Run `pnpm actions` to validate catalogue/handler/UI parity.");
  lines.push("- Run `pnpm scorecard:check` before release review.");
  lines.push(
    "- Treat `missing` and `partial` cells in the selected release slice as Go/No-Go inputs, then open or update tickets for accepted gaps."
  );
  lines.push("");
  lines.push("## Summary");
  lines.push("");
  lines.push("<!-- prettier-ignore-start -->");
  lines.push(
    "| Format | Actions | Handlers | UI-dispatched | MCP complete/partial/missing | Web complete/partial/missing | CLI complete/partial/missing |"
  );
  lines.push("| --- | ---: | ---: | ---: | ---: | ---: | ---: |");
  for (const f of scorecard.formats) {
    lines.push(
      `| ${f.format} | ${f.actions} | ${f.handlers} | ${f.uiDispatched} | ${surfaceSummary(f.summary.mcp)} | ${surfaceSummary(f.summary.web)} | ${surfaceSummary(f.summary.cli)} |`
    );
  }
  lines.push("");
  lines.push("## Matrix");
  lines.push("");
  lines.push("| Format | Capability | Operation | Command | MCP | Web | CLI | Evidence |");
  lines.push("| --- | --- | --- | --- | --- | --- | --- | --- |");
  for (const entry of scorecard.entries) {
    lines.push(
      `| ${entry.format} | ${escapeCell(entry.id)} | ${entry.operation} | ${escapeCell(entry.commandType ?? "-")} | ${entry.surfacesScore.mcp.status} | ${entry.surfacesScore.web.status} | ${entry.surfacesScore.cli.status} | ${escapeCell(compactEvidence(entry))} |`
    );
  }
  lines.push("<!-- prettier-ignore-end -->");
  lines.push("");
  lines.push("## Machine-Readable Companion");
  lines.push("");
  lines.push(
    "See [`cross-surface-parity-scorecard.json`](cross-surface-parity-scorecard.json) for per-surface notes and full evidence arrays."
  );
  lines.push("");
  return lines.join("\n");
}

function surfaceSummary(summary) {
  return `${summary.complete}/${summary.partial}/${summary.missing}`;
}

function compactEvidence(entry) {
  const bits = [];
  for (const surface of ["mcp", "web", "cli"]) {
    const score = entry.surfacesScore[surface];
    const first = score.evidence[0] ?? "none";
    bits.push(`${surface}:${first}`);
  }
  return bits.join("; ");
}

function escapeCell(value) {
  return String(value).replaceAll("|", "\\|").replaceAll("\n", " ");
}

function ensureParent(path) {
  const dir = dirname(join(ROOT, path));
  if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
}

function writeOrCheck(path, next) {
  const abs = join(ROOT, path);
  if (CHECK) {
    const current = existsSync(abs) ? readFileSync(abs, "utf8") : null;
    if (current !== next) {
      console.error(`${path} is stale. Run pnpm scorecard.`);
      return false;
    }
    return true;
  }
  ensureParent(path);
  writeFileSync(abs, next);
  return true;
}

function main() {
  const scorecard = buildScorecard();
  const json = `${JSON.stringify(scorecard, null, 2)}\n`;
  const md = renderMarkdown(scorecard);
  const jsonOk = writeOrCheck(JSON_OUT, json);
  const mdOk = writeOrCheck(MD_OUT, md);
  const ok = jsonOk && mdOk;
  if (!ok) return 1;
  if (CHECK) {
    console.log("surface scorecard: OK");
  } else {
    console.log(`surface scorecard: wrote ${MD_OUT} and ${JSON_OUT}`);
  }
  return 0;
}

process.exit(main());
