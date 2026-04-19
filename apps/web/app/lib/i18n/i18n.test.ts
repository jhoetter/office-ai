import { describe, expect, it } from "vitest";
import en from "./messages/en.json";
import de from "./messages/de.json";
import type { MessageNode } from "./types";

function collectKeys(node: MessageNode, prefix = ""): string[] {
  if (typeof node === "string") return prefix ? [prefix] : [];
  const out: string[] = [];
  for (const [key, value] of Object.entries(node)) {
    const next = prefix ? `${prefix}.${key}` : key;
    out.push(...collectKeys(value, next));
  }
  return out;
}

function collectPlaceholders(template: string): string[] {
  const matches = template.match(/\{(\w+)\}/g) ?? [];
  return matches.map((m) => m.slice(1, -1)).sort();
}

function flatten(node: MessageNode, prefix = ""): Record<string, string> {
  if (typeof node === "string") return prefix ? { [prefix]: node } : {};
  const out: Record<string, string> = {};
  for (const [key, value] of Object.entries(node)) {
    const next = prefix ? `${prefix}.${key}` : key;
    Object.assign(out, flatten(value, next));
  }
  return out;
}

describe("i18n catalogues", () => {
  it("English and German have the same set of keys", () => {
    const enKeys = new Set(collectKeys(en as MessageNode));
    const deKeys = new Set(collectKeys(de as MessageNode));
    const missingInDe = [...enKeys].filter((k) => !deKeys.has(k));
    const missingInEn = [...deKeys].filter((k) => !enKeys.has(k));
    expect(missingInDe, "keys present in en but missing in de").toEqual([]);
    expect(missingInEn, "keys present in de but missing in en").toEqual([]);
  });

  it("matching keys reuse the same {placeholders}", () => {
    const enFlat = flatten(en as MessageNode);
    const deFlat = flatten(de as MessageNode);
    const mismatches: Array<{ key: string; en: string[]; de: string[] }> = [];
    for (const [key, enValue] of Object.entries(enFlat)) {
      const deValue = deFlat[key];
      if (typeof deValue !== "string") continue;
      const enPlaceholders = collectPlaceholders(enValue);
      const dePlaceholders = collectPlaceholders(deValue);
      if (
        enPlaceholders.length !== dePlaceholders.length ||
        enPlaceholders.some((p, i) => p !== dePlaceholders[i])
      ) {
        mismatches.push({ key, en: enPlaceholders, de: dePlaceholders });
      }
    }
    expect(mismatches, "placeholder mismatch between locales").toEqual([]);
  });
});
