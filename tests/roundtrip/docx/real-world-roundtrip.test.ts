import { readFile, readdir } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml, sha256Hex } from "@officeai/core";
import { DocxAgent } from "@officeai/docx";

/**
 * Real-world fixture roundtrip suite (W1 / Batch P1.1).
 *
 * The fixtures under `fixtures/docx/real-world/` are produced by the
 * `docx` MIT npm library (see `scripts/generate-real-fixtures.mjs`) so
 * they ship the parts a real Word/LibreOffice/Google-Docs export would
 * ship: `word/styles.xml`, `word/numbering.xml`, headers/footers,
 * `word/comments.xml`, inline drawings.
 *
 * Invariants asserted here:
 *   1. Pure roundtrip (no mutation) keeps every part byte-identical.
 *   2. After a trivial edit (one inserted character at the start of the
 *      first paragraph), every part EXCEPT `word/document.xml` stays
 *      byte-identical. `document.xml` is re-emitted; the spec only
 *      requires structural equivalence on touched parts, which is
 *      covered by the unit tests in `packages/docx`.
 *
 * If a fixture surfaces a parser/serializer bug the test should record
 * it and skip with a pointer to the build log entry — do NOT fix the
 * bug here (W1 is the fixture/CI workstream, not the parser workstream).
 */

const FIXTURE_DIR = resolve(__dirname, "../../../fixtures/docx/real-world");
const DOC_XML_PART = "word/document.xml";

async function listFixtures(): Promise<string[]> {
  try {
    const entries = await readdir(FIXTURE_DIR);
    return entries.filter((f) => f.endsWith(".docx")).sort();
  } catch {
    return [];
  }
}

function partHashes(container: ooxml.OoxmlContainer): Record<string, string> {
  const out: Record<string, string> = {};
  for (const part of container.parts.keys()) {
    out[part] = sha256Hex(container.readBytes(part));
  }
  return out;
}

describe("DOCX real-world fixtures roundtrip", async () => {
  const fixtures = await listFixtures();

  if (fixtures.length === 0) {
    it("(no real-world fixtures present — run `pnpm fixtures-real`)", () => {
      expect(true).toBe(true);
    });
    return;
  }

  for (const name of fixtures) {
    const path = resolve(FIXTURE_DIR, name);

    it(`${name}: pure roundtrip preserves every part`, async () => {
      const buf = await readFile(path);
      const original = await ooxml.OoxmlContainer.load(buf);
      const before = partHashes(original);

      const agent = await DocxAgent.fromBuffer(buf);
      const out = await agent.exportFile();
      const reloaded = await ooxml.OoxmlContainer.load(out);
      const after = partHashes(reloaded);

      expect([...reloaded.parts.keys()].sort()).toEqual([...original.parts.keys()].sort());
      for (const part of Object.keys(before)) {
        expect(after[part], `pure-roundtrip part ${part}`).toBe(before[part]);
      }
    });

    it(`${name}: trivial edit only touches word/document.xml`, async () => {
      const buf = await readFile(path);
      const original = await ooxml.OoxmlContainer.load(buf);
      const before = partHashes(original);

      const agent = await DocxAgent.fromBuffer(buf);
      const snap = agent.getSnapshot();
      const firstParagraphIndex = snap.root.body.findIndex((b) => b.kind === "paragraph");
      // Every fixture is built around at least one paragraph; if it's
      // not, that's a content bug in the generator we want to surface.
      expect(firstParagraphIndex, `${name} contains no paragraph`).toBeGreaterThanOrEqual(0);

      await agent.applyCommand({
        type: "docx:insert-text",
        payload: {
          at: { paragraph: firstParagraphIndex, offset: 0 },
          text: "X",
        },
        source: "human",
      });

      const out = await agent.exportFile();
      const reloaded = await ooxml.OoxmlContainer.load(out);
      const after = partHashes(reloaded);

      expect([...reloaded.parts.keys()].sort()).toEqual([...original.parts.keys()].sort());
      for (const part of Object.keys(before)) {
        if (part === DOC_XML_PART) continue;
        expect(after[part], `untouched part ${part} should be byte-identical`).toBe(before[part]);
      }
      // Sanity: the edited part *did* change.
      expect(after[DOC_XML_PART]).not.toBe(before[DOC_XML_PART]);
    });
  }
});
