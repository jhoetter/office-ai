import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml, sha256Hex } from "@officeai/core";
import { parsePptx } from "../parser/parse.js";
import { serializePptx } from "./serialize.js";
import type { PptxSnapshot, TextShape } from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

/**
 * Mark one slide dirty without changing its model: this exercises the
 * "model → XML" path on that slide and proves that
 *   (a) the rebuilt slide bytes parse and
 *   (b) all other parts (presentation.xml, masters, layouts, theme, media,
 *       rels, content types, other slides) come back byte-identical.
 *
 * This is the strict gate from spec/pptx/acceptance-criteria.md §A2.
 */
describe("targeted-edit roundtrip", () => {
  it("rebuilds only the dirty slide; non-touched parts stay byte-identical", async () => {
    const path = join(FIXTURES_DIR.pathname, "04-multi-shape.pptx");
    const buf = await readFile(path);
    const snap = await parsePptx(buf);

    const slidePath = snap.root.slides[0].partPath;
    const dirtied: PptxSnapshot = {
      ...snap,
      dirty: { ...snap.dirty, slides: new Set([slidePath]) },
    };

    const out = await serializePptx(dirtied);
    const reload = await ooxml.OoxmlContainer.load(out);

    // Dirty slide is rebuilt — its bytes are not required to be identical,
    // only that it parses back into an equivalent model.
    expect(reload.has(slidePath)).toBe(true);

    // Every OTHER part must remain byte-identical.
    for (const partPath of snap.container.parts.keys()) {
      if (partPath === slidePath) continue;
      const before = sha256Hex(snap.container.readBytes(partPath));
      expect(reload.has(partPath), `part missing: ${partPath}`).toBe(true);
      const after = sha256Hex(reload.readBytes(partPath));
      expect(after, `${partPath} changed unexpectedly`).toBe(before);
    }
  });

  it("re-parses a dirty slide to an equivalent model (model-stable roundtrip)", async () => {
    const path = join(FIXTURES_DIR.pathname, "04-multi-shape.pptx");
    const buf = await readFile(path);
    const snap = await parsePptx(buf);
    const slidePath = snap.root.slides[0].partPath;
    const dirtied: PptxSnapshot = {
      ...snap,
      dirty: { ...snap.dirty, slides: new Set([slidePath]) },
    };
    const out = await serializePptx(dirtied);
    const snap2 = await parsePptx(out);

    const a = snap.root.slides[0];
    const b = snap2.root.slides[0];
    expect(b.shapes.length).toBe(a.shapes.length);
    for (let i = 0; i < a.shapes.length; i++) {
      const sa = a.shapes[i];
      const sb = b.shapes[i];
      expect(sb.kind).toBe(sa.kind);
      expect(sb.cNvPrId).toBe(sa.cNvPrId);
      expect(sb.name).toBe(sa.name);
      expect(sb.position).toEqual(sa.position);
      expect(sb.size).toEqual(sa.size);
    }
  });

  it("re-emits a text edit while leaving other parts byte-identical", async () => {
    const path = join(FIXTURES_DIR.pathname, "02-title-only.pptx");
    const buf = await readFile(path);
    const snap = await parsePptx(buf);

    // Mutate the first text shape's first run text.
    const slide = snap.root.slides[0];
    const textShape = slide.shapes.find((s): s is TextShape => s.kind === "text");
    expect(textShape).toBeDefined();
    if (!textShape) return;

    const newText = "Edited title 🎉";
    const newTextShape: TextShape = {
      ...textShape,
      txBody: {
        ...textShape.txBody,
        paragraphs: textShape.txBody.paragraphs.map((p, idx) =>
          idx === 0
            ? {
                ...p,
                runs: p.runs.map((r, ri) => (ri === 0 ? { ...r, text: newText } : r)),
              }
            : p
        ),
      },
    };

    const newSlide = {
      ...slide,
      shapes: slide.shapes.map((s) => (s.id === textShape.id ? newTextShape : s)),
    };

    const dirtied: PptxSnapshot = {
      ...snap,
      root: {
        ...snap.root,
        slides: snap.root.slides.map((s, i) => (i === 0 ? newSlide : s)),
      },
      dirty: { ...snap.dirty, slides: new Set([slide.partPath]) },
    };

    const out = await serializePptx(dirtied);
    const reload = await ooxml.OoxmlContainer.load(out);

    // Edited slide must contain the new text, and must NOT match its old hash.
    const newSlideBytes = reload.readBytes(slide.partPath);
    const newSlideText = new TextDecoder().decode(newSlideBytes);
    expect(newSlideText).toContain(newText);

    // Every other part must remain byte-identical.
    for (const partPath of snap.container.parts.keys()) {
      if (partPath === slide.partPath) continue;
      const before = sha256Hex(snap.container.readBytes(partPath));
      const after = sha256Hex(reload.readBytes(partPath));
      expect(after, `${partPath} changed unexpectedly`).toBe(before);
    }
  });
});
