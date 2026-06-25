import { readFileSync } from "node:fs";
import { resolve } from "node:path";

const ROOT = process.cwd();

const tokenFiles = [
  "packages/design-tokens/src/colors.ts",
  "packages/design-tokens/src/spacing.ts",
  "packages/design-tokens/src/typography.ts",
  "packages/design-tokens/src/tailwind-preset.ts",
  "apps/web/app/globals.css",
];

const legacyColorLiterals = [
  "#37352f",
  "#3b82f6",
  "#7c3aed",
  "#e9e9e7",
  "#fbfbfa",
  "#191919",
  "#202020",
  "#60a5fa",
  "#e57a2e",
  "#d84b3e",
  "#2f7d59",
];

const failures = [];

for (const file of tokenFiles) {
  const path = resolve(ROOT, file);
  const text = readFileSync(path, "utf8").toLowerCase();
  for (const literal of legacyColorLiterals) {
    if (text.includes(literal)) {
      failures.push(`${file}: found legacy local token ${literal}`);
    }
  }
}

const colors = readFileSync(resolve(ROOT, "packages/design-tokens/src/colors.ts"), "utf8");
const hardHex = colors.match(/#[0-9a-fA-F]{3,8}\b/g) ?? [];
if (hardHex.length > 0) {
  failures.push(`packages/design-tokens/src/colors.ts: hard-coded hex values ${hardHex.join(", ")}`);
}

if (!colors.includes("var(--sl-")) {
  failures.push("packages/design-tokens/src/colors.ts: expected sonaloop-design --sl-* aliases");
}

const globals = readFileSync(resolve(ROOT, "apps/web/app/globals.css"), "utf8");
for (const required of [
  '@import "sonaloop-design/fonts.css";',
  '@import "sonaloop-design/tokens.css";',
  '@import "sonaloop-design/components.css";',
  '@import "sonaloop-design/app.css";',
  "--background: var(--sl-bg);",
  "--foreground: var(--sl-ink);",
]) {
  if (!globals.includes(required)) {
    failures.push(`apps/web/app/globals.css: missing ${required}`);
  }
}

if (failures.length > 0) {
  console.error("Sonaloop token alias gate failed:");
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}

console.log("Sonaloop token alias gate: OK");
