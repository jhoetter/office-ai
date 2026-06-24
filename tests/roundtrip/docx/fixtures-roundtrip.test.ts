import { readFile, readdir } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml, sha256Hex } from "@officeai/core";
import { DocxAgent, parseDocx } from "@officeai/docx";
import { fixturePath, matrixFixtures } from "../../fixture-matrix.js";

/**
 * End-to-end round-trip test against the synthetic fixture corpus.
 *
 * Invariants:
 *  1. parseDocx → serializeDocx → load yields the same set of parts.
 *  2. Every part is byte-identical (we did not modify anything, so even
 *     edited parts round-trip when they were not changed).
 *  3. The agent can re-export and re-parse the file without error.
 */

const FIXTURE_DIR = resolve(__dirname, "../../../fixtures/docx/synthetic");

async function listFixtures(): Promise<Array<{ name: string; path: string }>> {
  const fromMatrix = matrixFixtures({
    format: "docx",
    origin: "synthetic",
    expectedBehavior: "noop-roundtrip",
  }).map((fixture) => ({ name: fixture.path.split("/").at(-1) ?? fixture.id, path: fixturePath(fixture) }));
  if (fromMatrix.length > 0) return fromMatrix;

  try {
    const entries = await readdir(FIXTURE_DIR);
    return entries
      .filter((f) => f.endsWith(".docx"))
      .sort()
      .map((name) => ({ name, path: resolve(FIXTURE_DIR, name) }));
  } catch {
    return [];
  }
}

describe("DOCX synthetic fixtures roundtrip", async () => {
  const fixtures = await listFixtures();

  if (fixtures.length === 0) {
    it("(no fixtures present — run `pnpm fixtures:docx`)", () => {
      expect(true).toBe(true);
    });
    return;
  }

  for (const { name, path } of fixtures) {
    it(`${name}: untouched parts are byte-identical after parse → serialize`, async () => {
      const buf = await readFile(path);

      const original = await ooxml.OoxmlContainer.load(buf);
      const originalHashes: Record<string, string> = {};
      for (const p of original.parts.keys()) {
        originalHashes[p] = sha256Hex(original.readBytes(p));
      }

      const snap = await parseDocx(buf);
      // No mutations applied. Serialize back and reload.
      const out = await (await DocxAgent.fromBuffer(buf)).exportFile();
      const reloaded = await ooxml.OoxmlContainer.load(out);

      const reloadedKeys = [...reloaded.parts.keys()].sort();
      const originalKeys = [...original.parts.keys()].sort();
      expect(reloadedKeys, "part list should be unchanged").toEqual(originalKeys);

      for (const p of originalKeys) {
        const after = sha256Hex(reloaded.readBytes(p));
        expect(after, `part ${p} should round-trip byte-for-byte`).toBe(originalHashes[p]);
      }
      expect(snap.format).toBe("docx");
    });

    it(`${name}: re-parses after agent export`, async () => {
      const buf = await readFile(path);
      const agent = await DocxAgent.fromBuffer(buf);
      const out = await agent.exportFile();
      const reparsed = await parseDocx(out);
      expect(reparsed.root.body.length).toBeGreaterThan(0);
    });
  }
});
