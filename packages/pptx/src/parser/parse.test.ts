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

  it("parses tables in p:graphicFrame as typed TableShape", async () => {
    const path = join(FIXTURES_DIR.pathname, "06-with-table.pptx");
    const buf = await readFile(path);
    const snap = await parsePptx(buf);
    const slide = snap.root.slides[0];
    const table = slide.shapes.find((s) => s.kind === "table");
    expect(table).toBeDefined();
    if (!table || table.kind !== "table") return;
    expect(table.graphicDataUri).toBe(
      "http://schemas.openxmlformats.org/drawingml/2006/table"
    );
    expect(table.columnWidths.length).toBe(3);
    for (const w of table.columnWidths) expect(w).toBeGreaterThan(0);
    expect(table.rows.length).toBeGreaterThanOrEqual(2);
    for (const row of table.rows) {
      expect(row.cells.length).toBe(table.columnWidths.length);
      for (const cell of row.cells) {
        expect(cell.txBody.paragraphs.length).toBeGreaterThanOrEqual(1);
      }
    }
    // First row, first cell should hold "Quarter" (header).
    const headerText = table.rows[0]!.cells[0]!.txBody.paragraphs
      .flatMap((p) => p.runs.filter((r) => !r.isLineBreak).map((r) => r.text))
      .join("");
    expect(headerText).toBe("Quarter");
    // Position/size are present (typed).
    expect(table.position?.xEmu).toBeGreaterThan(0);
    expect(table.size?.cxEmu).toBeGreaterThan(0);
  });

  it("parses charts in p:graphicFrame as typed ChartShape + ChartPart", async () => {
    const path = join(FIXTURES_DIR.pathname, "09-with-chart.pptx");
    const buf = await readFile(path);
    const snap = await parsePptx(buf);
    const slide = snap.root.slides[0];
    const chart = slide.shapes.find((s) => s.kind === "chart");
    expect(chart).toBeDefined();
    if (!chart || chart.kind !== "chart") return;
    expect(chart.graphicDataUri).toBe(
      "http://schemas.openxmlformats.org/drawingml/2006/chart"
    );
    expect(chart.chartRelId).toMatch(/^rId\d+$/);
    expect(chart.chartPartPath).toMatch(/^ppt\/charts\/chart\d+\.xml$/);
    expect(chart.position?.xEmu).toBeGreaterThan(0);
    expect(chart.size?.cxEmu).toBeGreaterThan(0);

    // Chart part is reachable + typed.
    const part = snap.root.charts.get(chart.chartPartPath);
    expect(part).toBeDefined();
    if (!part) return;
    expect(["bar", "line", "pie", "area", "unsupported"]).toContain(part.chartType);
    expect(part.series.length).toBeGreaterThanOrEqual(1);
    for (const s of part.series) {
      expect(s.values.length).toBeGreaterThan(0);
    }
    expect(part.categories.length).toBeGreaterThan(0);
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
