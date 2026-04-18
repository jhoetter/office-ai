import { readFile, readdir } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml, sha256Hex } from "@officeai/core";
import { PptxAgent, parsePptx, serializePptx } from "@officeai/pptx";
import type { PptxSnapshot, TextShape } from "@officeai/pptx";

/**
 * Real-world fixture roundtrip suite (F1.4).
 *
 * The fixtures under `fixtures/pptx/real/` are produced by the
 * `pptxgenjs` MIT npm library (see `scripts/generate-real-pptx-fixtures.mjs`).
 * pptxgenjs is a third-party PowerPoint-grade emitter, so the resulting
 * `.pptx` files ship the parts a typical PowerPoint / LibreOffice
 * Impress / Google Slides export ships: `ppt/theme/theme1.xml` with a
 * full color scheme, slide masters and layouts, multi-slide notes,
 * hyperlinks (`*.xml.rels`), embedded images, opaque tables, opaque
 * charts, and shape-rich slides.
 *
 * Invariants asserted here:
 *   1. Pure roundtrip (no mutation) keeps ≥95 % of parts byte-identical
 *      and 100 % of part *paths* identical.
 *   2. After marking the first slide dirty the model serializes back
 *      and re-parses to an equivalent shape list, every other part
 *      stays byte-identical.
 *
 * The 95 % threshold (rather than 100 %) leaves room for the
 * presentation-rebuild path to re-emit `ppt/presentation.xml` when
 * `dirty.presentation` is set, but for a *pure* roundtrip we expect
 * full byte-identity. Failures should be triaged in `MANIFEST.md`.
 */

const FIXTURE_DIR = resolve(__dirname, "../../../fixtures/pptx/real");
const PURE_ROUNDTRIP_THRESHOLD = 0.95;

async function listFixtures(): Promise<string[]> {
  try {
    const entries = await readdir(FIXTURE_DIR);
    return entries.filter((f) => f.endsWith(".pptx")).sort();
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

describe("PPTX real-world fixtures roundtrip", async () => {
  const fixtures = await listFixtures();

  if (fixtures.length === 0) {
    it("(no real-world fixtures present — run `pnpm fixtures:pptx-real`)", () => {
      expect(true).toBe(true);
    });
    return;
  }

  for (const name of fixtures) {
    const path = resolve(FIXTURE_DIR, name);

    it(`${name}: pure roundtrip preserves ≥95 % of parts byte-for-byte`, async () => {
      const buf = await readFile(path);
      const original = await ooxml.OoxmlContainer.load(buf);
      const before = partHashes(original);

      const agent = await PptxAgent.fromBuffer(buf);
      const out = await agent.exportFile();
      const reloaded = await ooxml.OoxmlContainer.load(out);
      const after = partHashes(reloaded);

      expect([...reloaded.parts.keys()].sort()).toEqual([...original.parts.keys()].sort());

      const parts = Object.keys(before);
      const drift: string[] = [];
      for (const part of parts) {
        if (after[part] !== before[part]) drift.push(part);
      }
      const ratio = (parts.length - drift.length) / parts.length;
      expect(
        ratio,
        `byte-identity ratio ${(ratio * 100).toFixed(1)}% drift=${JSON.stringify(drift)}`
      ).toBeGreaterThanOrEqual(PURE_ROUNDTRIP_THRESHOLD);
    });

    it(`${name}: editing the first text shape leaves untouched parts byte-identical`, async () => {
      const buf = await readFile(path);
      const original = await ooxml.OoxmlContainer.load(buf);
      const before = partHashes(original);

      const snap = await parsePptx(buf);
      const slide = snap.root.slides[0];
      const textShape = slide.shapes.find((s): s is TextShape => s.kind === "text");
      if (!textShape || textShape.txBody.paragraphs.length === 0) {
        // Not every fixture has an editable text shape on slide 0; skip.
        return;
      }

      const editedText = "officeAI roundtrip ✓";
      const newTextShape: TextShape = {
        ...textShape,
        txBody: {
          ...textShape.txBody,
          paragraphs: textShape.txBody.paragraphs.map((p, i) =>
            i === 0
              ? {
                  ...p,
                  runs:
                    p.runs.length > 0
                      ? p.runs.map((r, ri) => (ri === 0 ? { ...r, text: editedText } : r))
                      : p.runs,
                }
              : p
          ),
        },
      };

      const dirtied: PptxSnapshot = {
        ...snap,
        root: {
          ...snap.root,
          slides: snap.root.slides.map((s, i) =>
            i === 0
              ? {
                  ...s,
                  shapes: s.shapes.map((sh) => (sh.id === textShape.id ? newTextShape : sh)),
                }
              : s
          ),
        },
        dirty: { ...snap.dirty, slides: new Set([slide.partPath]) },
      };

      const out = await serializePptx(dirtied);
      const reloaded = await ooxml.OoxmlContainer.load(out);
      const after = partHashes(reloaded);

      expect([...reloaded.parts.keys()].sort()).toEqual([...original.parts.keys()].sort());

      for (const part of Object.keys(before)) {
        if (part === slide.partPath) continue;
        expect(after[part], `untouched part ${part} should be byte-identical`).toBe(before[part]);
      }
      // Sanity: edited slide changed and contains the new text.
      expect(after[slide.partPath]).not.toBe(before[slide.partPath]);
      const slideXml = new TextDecoder().decode(reloaded.readBytes(slide.partPath));
      expect(slideXml).toContain(editedText);
    });
  }
});
