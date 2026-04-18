import { readFile, readdir } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";

const SRC_DIR = new URL("../src", import.meta.url).pathname;

const FORBIDDEN_HEADLESS_IMPORTS = [
  "react",
  "react-dom",
  "react/jsx-runtime",
  "next",
  "next/router",
  "next/navigation",
  "tailwindcss",
];

const FORBIDDEN_HEADLESS_GLOBALS = [
  /(?<![A-Za-z0-9_."'`])document\.[a-zA-Z]/,
  /(?<![A-Za-z0-9_."'`])window\.[a-zA-Z]/,
  /(?<![A-Za-z0-9_."'`])navigator\.[a-zA-Z]/,
  /(?<![A-Za-z0-9_."'`])location\.[a-zA-Z]/,
];

function stripStringsAndComments(src: string): string {
  // Remove block comments, line comments, and string/template literals so
  // we don't match `"document.xml"` or `// document.foo` as DOM access.
  let out = src.replace(/\/\*[\s\S]*?\*\//g, "");
  out = out.replace(/(^|[^:])\/\/.*$/gm, "$1");
  out = out.replace(/"(?:[^"\\]|\\.)*"/g, '""');
  out = out.replace(/'(?:[^'\\]|\\.)*'/g, "''");
  out = out.replace(/`(?:[^`\\]|\\.)*`/g, "``");
  return out;
}

const HEADLESS_DIRS = [
  "model",
  "parser",
  "serializer",
  "commands",
  "agent",
  "renderer/layout",
  "renderer/svg",
];

async function* walk(dir: string): AsyncGenerator<string> {
  const entries = await readdir(dir, { withFileTypes: true });
  for (const e of entries) {
    if (e.name.startsWith(".")) continue;
    const p = join(dir, e.name);
    if (e.isDirectory()) {
      yield* walk(p);
    } else if (e.isFile() && (p.endsWith(".ts") || p.endsWith(".tsx"))) {
      yield p;
    }
  }
}

describe("headless-purity invariant", () => {
  it("has no DOM/React imports in model/parser/serializer/commands/agent", async () => {
    const offenders: string[] = [];
    for (const subdir of HEADLESS_DIRS) {
      const root = join(SRC_DIR, subdir);
      try {
        await readdir(root);
      } catch {
        continue;
      }
      for await (const file of walk(root)) {
        if (file.endsWith(".test.ts") || file.endsWith(".test.tsx")) continue;
        const content = await readFile(file, "utf8");
        for (const dep of FORBIDDEN_HEADLESS_IMPORTS) {
          const re = new RegExp(`from\\s+["']${escapeRe(dep)}["']`);
          if (re.test(content)) {
            offenders.push(`${file}: forbidden import of ${dep}`);
          }
        }
        const code = stripStringsAndComments(content);
        for (const g of FORBIDDEN_HEADLESS_GLOBALS) {
          if (g.test(code)) {
            offenders.push(`${file}: forbidden global use ${g}`);
          }
        }
      }
    }
    expect(offenders, offenders.join("\n")).toEqual([]);
  });
});

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
