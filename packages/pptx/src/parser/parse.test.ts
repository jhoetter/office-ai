import { readFile, readdir } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { parsePptx } from "./parse.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function listFixtures(): Promise<string[]> {
  const dir = FIXTURES_DIR.pathname;
  const entries = await readdir(dir);
  return entries.filter((e) => e.endsWith(".pptx")).sort();
}

describe("parsePptx", () => {
  it("parses every synthetic fixture", async () => {
    const fixtures = await listFixtures();
    expect(fixtures.length).toBeGreaterThan(0);

    for (const name of fixtures) {
      const path = join(FIXTURES_DIR.pathname, name);
      const buf = await readFile(path);
      const snap = await parsePptx(buf);
      expect(snap.format).toBe("pptx");
      expect(snap.root.slides.length).toBeGreaterThan(0);
      // Slide ids must be ≥ 256 and unique.
      const ids = new Set<number>();
      for (const s of snap.root.slides) {
        expect(s.slideId).toBeGreaterThanOrEqual(256);
        expect(ids.has(s.slideId)).toBe(false);
        ids.add(s.slideId);
        expect(s.relId).toMatch(/^rId\d+$/);
      }
      // Slide size present.
      expect(snap.root.slideSize.cxEmu).toBeGreaterThan(0);
      expect(snap.root.slideSize.cyEmu).toBeGreaterThan(0);
      // Container is attached.
      expect(snap.container.has("ppt/presentation.xml")).toBe(true);
      // partHashes covers every part.
      for (const p of snap.container.parts.keys()) {
        expect(snap.partHashes[p]).toBeDefined();
      }
    }
  });

  it("resolves picture media paths", async () => {
    const path = join(FIXTURES_DIR.pathname, "05-with-image.pptx");
    const buf = await readFile(path);
    const snap = await parsePptx(buf);
    const slide = snap.root.slides[0];
    const pic = slide.shapes.find((s) => s.kind === "pic");
    expect(pic).toBeDefined();
    if (pic && pic.kind === "pic") {
      expect(pic.mediaRelId).toMatch(/^rId\d+$/);
      expect(pic.mediaPartPath).toMatch(/^ppt\/media\//);
      expect(snap.root.media.has(pic.mediaPartPath)).toBe(true);
    }
  });

  it("captures multiple slides in order", async () => {
    const path = join(FIXTURES_DIR.pathname, "07-multi-slide.pptx");
    const buf = await readFile(path);
    const snap = await parsePptx(buf);
    expect(snap.root.slides.length).toBeGreaterThanOrEqual(3);
    // slideId monotonically non-decreasing in declared order? Not required by
    // OOXML, but our generator emits ascending ids.
    const ids = snap.root.slides.map((s) => s.slideId);
    const sorted = [...ids].sort((a, b) => a - b);
    expect(ids).toEqual(sorted);
  });
});
