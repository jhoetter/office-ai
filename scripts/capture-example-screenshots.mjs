#!/usr/bin/env node
import { mkdirSync, readFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(fileURLToPath(import.meta.url), "..", "..");

function parseArgs(argv) {
  const out = { example: undefined, vars: {}, baseUrl: undefined, dryRun: false };
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--base-url") out.baseUrl = argv[++i];
    else if (arg === "--document-id") out.vars.documentId = argv[++i];
    else if (arg === "--session-id") out.vars.sessionId = argv[++i];
    else if (arg === "--dry-run") out.dryRun = true;
    else if (!out.example) out.example = arg;
    else throw new Error(`Unexpected argument: ${arg}`);
  }
  if (!out.example)
    throw new Error(
      "Usage: node scripts/capture-example-screenshots.mjs <example-id> [--base-url URL] [--document-id ID] [--session-id ID] [--dry-run]"
    );
  return out;
}

function readJson(path) {
  return JSON.parse(readFileSync(path, "utf8"));
}

function substitute(input, vars) {
  return input.replace(/\$\{([A-Za-z0-9_]+)\}/g, (_match, key) => {
    const value = vars[key] ?? process.env[`OFFICEAI_EXAMPLE_${key.toUpperCase()}`];
    if (!value) throw new Error(`Missing value for screenshot variable \${${key}}`);
    return value;
  });
}

async function loadChromium() {
  try {
    const mod = await import("playwright");
    return mod.chromium;
  } catch {
    try {
      const mod = await import("@playwright/test");
      return mod.chromium;
    } catch {
      throw new Error(
        "Playwright is not available. Run `pnpm --filter @officeai/web exec playwright install chromium` after installing dependencies."
      );
    }
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const exampleDir = join(ROOT, "examples", args.example);
  const manifest = readJson(join(exampleDir, "manifest.json"));
  if (!manifest.web?.screenshotPlan) throw new Error(`${args.example} does not declare a web.screenshotPlan`);
  const plan = readJson(join(exampleDir, manifest.web.screenshotPlan));
  const baseUrl = args.baseUrl ?? plan.baseUrl ?? "http://localhost:3100";
  const outputDir = join(ROOT, plan.outputDir);

  const targets = plan.targets.map((target) => {
    const path = substitute(target.path, args.vars);
    return {
      name: target.name,
      url: new URL(path, baseUrl).toString(),
      output: join(outputDir, target.output),
      waitFor: target.waitFor ?? "load",
    };
  });

  if (args.dryRun) {
    for (const target of targets) console.log(`${target.name}: ${target.url} -> ${target.output}`);
    return;
  }

  mkdirSync(outputDir, { recursive: true });
  const chromium = await loadChromium();
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage({ viewport: { width: 1440, height: 960 } });
    for (const target of targets) {
      await page.goto(target.url, { waitUntil: target.waitFor });
      await page.screenshot({ path: target.output, fullPage: true });
      console.log(`captured ${target.name}: ${target.output}`);
    }
  } finally {
    await browser.close();
  }
}

main().catch((err) => {
  console.error(err instanceof Error ? err.message : String(err));
  process.exit(1);
});
