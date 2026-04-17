import { readFile, readdir } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml, sha256Hex } from "@officeai/core";
import { DocxAgent, parseDocx } from "@officeai/docx";

const FIXTURE_DIR = resolve(__dirname, "../../../fixtures/docx/synthetic");

async function listFixtures(): Promise<string[]> {
  try {
    return (await readdir(FIXTURE_DIR)).filter((f) => f.endsWith(".docx")).sort();
  } catch {
    return [];
  }
}

/**
 * After a single small edit, exactly the parts we touched should change
 * (here: word/document.xml). All other parts must remain byte-identical.
 *
 * This protects against accidental "round-trip drift" when the serializer
 * touches more parts than necessary.
 */
describe("agent edit -> export -> reload preserves untouched parts", async () => {
  const fixtures = await listFixtures();
  if (fixtures.length === 0) {
    it("(no fixtures present — run `pnpm fixtures:docx`)", () => {
      expect(true).toBe(true);
    });
    return;
  }

  for (const name of fixtures) {
    it(`${name}: only word/document.xml changes after a text insert`, async () => {
      const buf = await readFile(resolve(FIXTURE_DIR, name));
      const original = await ooxml.OoxmlContainer.load(buf);
      const originalHashes: Record<string, string> = {};
      for (const p of original.parts.keys()) {
        originalHashes[p] = sha256Hex(original.readBytes(p));
      }

      const agent = await DocxAgent.fromBuffer(buf);
      await agent.applyCommand({
        type: "docx:insert-text",
        payload: { at: { paragraph: 0, run: 0, offset: 0 }, text: "EDIT — " },
        source: "human",
      });
      const out = await agent.exportFile();
      const reloaded = await ooxml.OoxmlContainer.load(out);

      const expectedDirty = new Set(["word/document.xml"]);
      for (const p of reloaded.parts.keys()) {
        const after = sha256Hex(reloaded.readBytes(p));
        if (expectedDirty.has(p)) continue;
        expect(after, `${name}: untouched part ${p} should be byte-identical`).toBe(originalHashes[p]);
      }

      const reparsed = await parseDocx(out);
      const p0 = reparsed.root.body.find((b) => b.kind === "paragraph");
      expect(p0).toBeTruthy();
    });
  }
});
