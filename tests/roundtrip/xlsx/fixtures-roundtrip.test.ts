import { readFile, readdir } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml, sha256Hex } from "@officeai/core";
import { parseXlsx, serializeXlsx } from "@officeai/xlsx";

/**
 * End-to-end round-trip test against the synthetic XLSX fixture corpus.
 *
 * Invariants:
 *  1. parseXlsx → serializeXlsx → load yields the same set of parts.
 *  2. Every part is byte-content-identical (no mutations applied, so
 *     the dirty set is empty and every part is re-emitted verbatim).
 *  3. Re-parsing the output produces a structurally equivalent
 *     workbook (same sheet count, same names, same tab order).
 */

const FIXTURE_DIR = resolve(__dirname, "../../../fixtures/xlsx/synthetic");

async function listFixtures(): Promise<string[]> {
  try {
    const entries = await readdir(FIXTURE_DIR);
    return entries.filter((f) => f.endsWith(".xlsx")).sort();
  } catch {
    return [];
  }
}

describe("XLSX synthetic fixtures roundtrip", async () => {
  const fixtures = await listFixtures();

  if (fixtures.length === 0) {
    it("(no fixtures present — run `pnpm fixtures:xlsx`)", () => {
      expect(true).toBe(true);
    });
    return;
  }

  for (const name of fixtures) {
    it(`${name}: untouched parts are byte-identical after parse → serialize`, async () => {
      const path = resolve(FIXTURE_DIR, name);
      const buf = await readFile(path);

      const original = await ooxml.OoxmlContainer.load(buf);
      const originalHashes: Record<string, string> = {};
      for (const p of original.parts.keys()) {
        originalHashes[p] = sha256Hex(original.readBytes(p));
      }

      const snap = await parseXlsx(buf);
      const out = await serializeXlsx(snap);
      const reloaded = await ooxml.OoxmlContainer.load(out);

      const reloadedKeys = [...reloaded.parts.keys()].sort();
      const originalKeys = [...original.parts.keys()].sort();
      expect(reloadedKeys, "part list should be unchanged").toEqual(originalKeys);

      for (const p of originalKeys) {
        const after = sha256Hex(reloaded.readBytes(p));
        expect(after, `part ${p} should round-trip byte-for-byte`).toBe(originalHashes[p]);
      }
      expect(snap.format).toBe("xlsx");
    });

    it(`${name}: re-parses after serialize, sheet list is structurally equal`, async () => {
      const path = resolve(FIXTURE_DIR, name);
      const buf = await readFile(path);
      const original = await parseXlsx(buf);
      const out = await serializeXlsx(original);
      const reparsed = await parseXlsx(new Uint8Array(out));
      expect(reparsed.root.sheets.map((s) => s.name)).toEqual(original.root.sheets.map((s) => s.name));
      expect(reparsed.root.sheets.map((s) => s.partPath)).toEqual(
        original.root.sheets.map((s) => s.partPath)
      );
      expect(reparsed.root.sheets.map((s) => s.kind)).toEqual(original.root.sheets.map((s) => s.kind));
    });
  }
});
