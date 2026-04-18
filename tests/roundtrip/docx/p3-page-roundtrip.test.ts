import { readFile, readdir } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml, sha256Hex } from "@officeai/core";
import { DocxAgent } from "@officeai/docx";

/**
 * P3.7 / W26 — page-aware round-trip invariants for the real-world
 * fixtures. Sits alongside `real-world-roundtrip.test.ts`; the
 * existing test guarantees byte-identical preservation of every
 * untouched part after a no-op load → save. This file extends that
 * with the four mutations introduced in P3.4 / P3.5 and asserts the
 * isolation property still holds:
 *
 *   1. Inserting a hard page break in the body re-emits only
 *      `word/document.xml`. Headers, footers, styles, numbering,
 *      comments, and every other part stay byte-identical.
 *   2. Inserting a `<w:fldSimple w:instr=" PAGE "/>` field into the
 *      first header paragraph re-emits only that header part (when
 *      one exists). Body and other parts stay byte-identical.
 *   3. Toggling `differentFirstPage` on the trailing implicit
 *      section re-emits only `word/document.xml` (the typed
 *      SectionProperties model lives inside the body's `<w:sectPr>`).
 *   4. Inserting a continuous section break before paragraph 1
 *      re-emits only `word/document.xml`.
 *
 * These are the editor mutations P3 surfaces as user-facing — every
 * one of them is also wired into MCP via `docx_apply_command`, so the
 * isolation property here doubles as the guarantee MCP tools won't
 * accidentally rewrite unrelated parts.
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

describe("DOCX P3 page-aware round-trip invariants", async () => {
  const fixtures = await listFixtures();

  if (fixtures.length === 0) {
    it("(no real-world fixtures present — run `pnpm fixtures-real`)", () => {
      expect(true).toBe(true);
    });
    return;
  }

  for (const name of fixtures) {
    const path = resolve(FIXTURE_DIR, name);

    it(`${name}: insert-page-break re-emits only word/document.xml`, async () => {
      const buf = await readFile(path);
      const original = await ooxml.OoxmlContainer.load(buf);
      const before = partHashes(original);

      const agent = await DocxAgent.fromBuffer(buf);
      const snap = agent.getSnapshot();
      const firstParaIdx = snap.root.body.findIndex((b) => b.kind === "paragraph");
      if (firstParaIdx < 0) return; // Fixture without body paragraphs — skip
      const firstPara = snap.root.body[firstParaIdx];
      if (firstPara.kind !== "paragraph") return;

      const m = await agent.applyCommand({
        type: "docx:insert-page-break",
        payload: { paragraphId: firstPara.id, offset: 0 },
        source: "human",
      });
      expect(m.status).toBe("approved");

      const out = await agent.exportFile();
      const reloaded = await ooxml.OoxmlContainer.load(out);
      const after = partHashes(reloaded);
      for (const part of Object.keys(before)) {
        if (part === DOC_XML_PART) continue;
        expect(after[part], `${part} should be untouched by insert-page-break`).toBe(before[part]);
      }
      expect(after[DOC_XML_PART]).not.toBe(before[DOC_XML_PART]);
    });

    it(`${name}: set-section-different-first re-emits only word/document.xml`, async () => {
      const buf = await readFile(path);
      const original = await ooxml.OoxmlContainer.load(buf);
      const before = partHashes(original);

      const agent = await DocxAgent.fromBuffer(buf);
      const m = await agent.applyCommand({
        type: "docx:set-section-different-first",
        payload: { paragraphIndex: 0, enabled: true },
        source: "human",
      });
      // Either approved (section was found) or rejected with
      // unknown-target (fixture has no section). Both are fine — we
      // just want to make sure no body-isolation invariant is
      // violated when the command does succeed.
      if (m.status !== "approved") return;

      const out = await agent.exportFile();
      const reloaded = await ooxml.OoxmlContainer.load(out);
      const after = partHashes(reloaded);
      for (const part of Object.keys(before)) {
        if (part === DOC_XML_PART) continue;
        expect(after[part], `${part} should be untouched by titlePg toggle`).toBe(before[part]);
      }
    });

    it(`${name}: insert-section-break re-emits only word/document.xml`, async () => {
      const buf = await readFile(path);
      const original = await ooxml.OoxmlContainer.load(buf);
      const before = partHashes(original);

      const agent = await DocxAgent.fromBuffer(buf);
      const m = await agent.applyCommand({
        type: "docx:insert-section-break",
        payload: { paragraphIndex: 0, type: "continuous" },
        source: "human",
      });
      if (m.status !== "approved") return;

      const out = await agent.exportFile();
      const reloaded = await ooxml.OoxmlContainer.load(out);
      const after = partHashes(reloaded);
      for (const part of Object.keys(before)) {
        if (part === DOC_XML_PART) continue;
        expect(after[part], `${part} should be untouched by insert-section-break`).toBe(before[part]);
      }
    });

    it(`${name}: insert-page-number into first header re-emits only that header part`, async () => {
      const buf = await readFile(path);
      const original = await ooxml.OoxmlContainer.load(buf);
      const before = partHashes(original);

      const agent = await DocxAgent.fromBuffer(buf);
      const snap = agent.getSnapshot();
      const header = snap.root.headersAndFooters.find((p) => p.kind === "header");
      if (!header) return; // Fixture without headers — skip
      const firstPara = header.body.find((b) => b.kind === "paragraph");
      if (!firstPara || firstPara.kind !== "paragraph") return;

      const m = await agent.applyCommand({
        type: "docx:insert-page-number",
        payload: { paragraphId: firstPara.id, offset: 0 },
        source: "human",
      });
      expect(m.status).toBe("approved");

      const out = await agent.exportFile();
      const reloaded = await ooxml.OoxmlContainer.load(out);
      const after = partHashes(reloaded);
      for (const part of Object.keys(before)) {
        if (part === header.partPath) continue;
        // word/document.xml shouldn't change — page numbers live in
        // the header part only. If it does change something is wrong.
        expect(after[part], `${part} should be untouched by header page-number insert`).toBe(before[part]);
      }
      expect(after[header.partPath]).not.toBe(before[header.partPath]);
    });

    it(`${name}: page chunker output is stable across pure round-trip`, async () => {
      const buf = await readFile(path);
      const a1 = await DocxAgent.fromBuffer(buf);
      const before = a1.getPages().map((p) => ({
        pageNumber: p.pageNumber,
        startBlockIndex: p.startBlockIndex,
        endBlockIndex: p.endBlockIndex,
        trigger: p.trigger,
      }));
      const out = await a1.exportFile();
      const a2 = await DocxAgent.fromBuffer(out);
      const after = a2.getPages().map((p) => ({
        pageNumber: p.pageNumber,
        startBlockIndex: p.startBlockIndex,
        endBlockIndex: p.endBlockIndex,
        trigger: p.trigger,
      }));
      expect(after).toEqual(before);
    });
  }
});
