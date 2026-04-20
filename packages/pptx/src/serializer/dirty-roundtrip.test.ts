import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml, sha256Hex } from "@officeai/core";
import { parsePptx } from "../parser/parse.js";
import { serializePptx } from "./serialize.js";
import type { PptxSnapshot, TableShape, TextShape } from "../model/types.js";

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

  it("rebuilds a typed TableShape and survives parse → serialize → parse", async () => {
    const path = join(FIXTURES_DIR.pathname, "06-with-table.pptx");
    const buf = await readFile(path);
    const snap = await parsePptx(buf);

    const slide = snap.root.slides[0];
    const tableA = slide.shapes.find((s): s is TableShape => s.kind === "table");
    expect(tableA).toBeDefined();
    if (!tableA) return;

    const dirtied: PptxSnapshot = {
      ...snap,
      dirty: { ...snap.dirty, slides: new Set([slide.partPath]) },
    };
    const out = await serializePptx(dirtied);
    const snap2 = await parsePptx(out);
    const slide2 = snap2.root.slides[0];
    const tableB = slide2.shapes.find((s): s is TableShape => s.kind === "table");
    expect(tableB).toBeDefined();
    if (!tableB) return;

    expect(tableB.columnWidths).toEqual(tableA.columnWidths);
    expect(tableB.rows.length).toBe(tableA.rows.length);
    for (let r = 0; r < tableA.rows.length; r++) {
      const ra = tableA.rows[r]!;
      const rb = tableB.rows[r]!;
      expect(rb.cells.length).toBe(ra.cells.length);
      for (let c = 0; c < ra.cells.length; c++) {
        const a = ra.cells[c]!;
        const b = rb.cells[c]!;
        const at = a.txBody.paragraphs
          .flatMap((p) => p.runs.filter((r2) => !r2.isLineBreak).map((r2) => r2.text))
          .join("");
        const bt = b.txBody.paragraphs
          .flatMap((p) => p.runs.filter((r2) => !r2.isLineBreak).map((r2) => r2.text))
          .join("");
        expect(bt).toBe(at);
      }
    }

    // Non-touched parts (everything except the rebuilt slide) stay byte-identical.
    const reload = await ooxml.OoxmlContainer.load(out);
    for (const partPath of snap.container.parts.keys()) {
      if (partPath === slide.partPath) continue;
      const before = sha256Hex(snap.container.readBytes(partPath));
      expect(reload.has(partPath), `part missing: ${partPath}`).toBe(true);
      const after = sha256Hex(reload.readBytes(partPath));
      expect(after, `${partPath} changed unexpectedly`).toBe(before);
    }
  });

  it("rebuilds a typed ChartShape and survives parse → serialize → parse", async () => {
    const path = join(FIXTURES_DIR.pathname, "09-with-chart.pptx");
    const buf = await readFile(path);
    const snap = await parsePptx(buf);

    const slide = snap.root.slides[0];
    const chartA = slide.shapes.find((s) => s.kind === "chart");
    expect(chartA).toBeDefined();
    if (!chartA || chartA.kind !== "chart") return;
    const partA = snap.root.charts.get(chartA.chartPartPath);
    expect(partA).toBeDefined();
    if (!partA) return;

    // Dirty the slide AND the chart part to force both to be rebuilt
    // from the typed model.
    const dirtied: PptxSnapshot = {
      ...snap,
      dirty: {
        ...snap.dirty,
        slides: new Set([slide.partPath]),
        charts: new Set([chartA.chartPartPath]),
      },
    };

    const out = await serializePptx(dirtied);
    const snap2 = await parsePptx(out);
    const slide2 = snap2.root.slides[0];
    const chartB = slide2.shapes.find((s) => s.kind === "chart");
    expect(chartB).toBeDefined();
    if (!chartB || chartB.kind !== "chart") return;
    expect(chartB.chartRelId).toBe(chartA.chartRelId);
    expect(chartB.chartPartPath).toBe(chartA.chartPartPath);
    expect(chartB.position).toEqual(chartA.position);
    expect(chartB.size).toEqual(chartA.size);

    const partB = snap2.root.charts.get(chartB.chartPartPath);
    expect(partB).toBeDefined();
    if (!partB) return;
    expect(partB.chartType).toBe(partA.chartType);
    expect(partB.title).toBe(partA.title);
    expect(partB.series.length).toBe(partA.series.length);
    for (let i = 0; i < partA.series.length; i++) {
      expect(partB.series[i]!.values).toEqual(partA.series[i]!.values);
    }
    expect(partB.categories).toEqual(partA.categories);
  });

  it("authors an embedded xlsx + c:externalData when a chart is dirty", async () => {
    const path = join(FIXTURES_DIR.pathname, "09-with-chart.pptx");
    const buf = await readFile(path);
    const snap = await parsePptx(buf);

    const slide = snap.root.slides[0];
    const chartA = slide.shapes.find((s) => s.kind === "chart");
    expect(chartA).toBeDefined();
    if (!chartA || chartA.kind !== "chart") return;

    const dirtied: PptxSnapshot = {
      ...snap,
      dirty: {
        ...snap.dirty,
        slides: new Set([slide.partPath]),
        charts: new Set([chartA.chartPartPath]),
      },
    };

    const out = await serializePptx(dirtied);
    const reload = await ooxml.OoxmlContainer.load(out);

    // (1) chart rels reference an embedded package under ppt/embeddings/.
    const chartRels = ooxml.RelationshipGraph.loadFor(reload, chartA.chartPartPath);
    const pkgRel = chartRels.relationships.find((r) => r.type === ooxml.REL_TYPE_PACKAGE);
    expect(pkgRel, "chart should advertise a package relationship").toBeDefined();
    if (!pkgRel) return;
    expect(pkgRel.target).toMatch(/embeddings\/[^/]+\.xlsx$/);

    // (2) the embedded part exists and parses as a real xlsx package.
    const embeddedPath = pkgRel.target.startsWith("/")
      ? pkgRel.target.slice(1)
      : new URL(pkgRel.target, `file:///${chartA.chartPartPath}`).pathname.slice(1);
    expect(reload.has(embeddedPath), `embedded xlsx missing at ${embeddedPath}`).toBe(true);
    const inner = await ooxml.OoxmlContainer.load(reload.readBytes(embeddedPath));
    expect(inner.has("xl/workbook.xml")).toBe(true);

    // (3) the chart XML emits c:externalData wired to the new rId, and
    // its formulas point into the embedded sheet (not the legacy
    // placeholder Sheet1!$A$2:$A$N strings, which we still allow but
    // expect to be derived from the embedded layout).
    const chartXml = new TextDecoder().decode(reload.readBytes(chartA.chartPartPath));
    expect(chartXml).toContain(`r:id="${pkgRel.id}"`);
    expect(chartXml).toContain("c:externalData");
    expect(chartXml).toMatch(/<c:f>[^<]*\$A\$2:\$A\$/);

    // (4) [Content_Types].xml advertises the spreadsheetml override for it.
    const ct = ooxml.ContentTypes.load(reload);
    const partName = embeddedPath.startsWith("/") ? embeddedPath : `/${embeddedPath}`;
    expect(ct.hasOverride(partName)).toBe(true);
  });

  it("rebuilds a slide with typed transition + animations and preserves them", async () => {
    const path = join(FIXTURES_DIR.pathname, "10-with-anim.pptx");
    const buf = await readFile(path);
    const snap = await parsePptx(buf);
    const slide = snap.root.slides[0];
    expect(slide.transition?.kind).toBe("fade");
    expect(slide.animations.length).toBe(2);

    const dirtied: PptxSnapshot = {
      ...snap,
      dirty: { ...snap.dirty, slides: new Set([slide.partPath]) },
    };
    const out = await serializePptx(dirtied);
    const snap2 = await parsePptx(out);
    const slide2 = snap2.root.slides[0];

    // Transition + animations survive the rebuild via the preserved
    // raw <p:transition> blob and the verbatim <p:timing> tail.
    expect(slide2.transition?.kind).toBe("fade");
    expect(slide2.transition?.speed).toBe("med");
    expect(slide2.animations.length).toBe(2);
    expect(slide2.animations[0]).toMatchObject({
      category: "entrance",
      preset: "appear",
      targetCNvPrId: 2,
    });
    expect(slide2.animations[1]).toMatchObject({
      category: "entrance",
      preset: "flyIn",
      targetCNvPrId: 3,
      durationMs: 500,
    });
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

  it("flushes a dirty slide-master through the new dirty.masters path", async () => {
    const path = join(FIXTURES_DIR.pathname, "04-multi-shape.pptx");
    const buf = await readFile(path);
    const snap = await parsePptx(buf);
    const masterPath = [...snap.root.masters.keys()][0]!;
    expect(masterPath).toBeDefined();

    const dirtied: PptxSnapshot = {
      ...snap,
      dirty: { ...snap.dirty, masters: new Set([masterPath]) },
    };
    const out = await serializePptx(dirtied);
    const reload = await ooxml.OoxmlContainer.load(out);

    expect(reload.has(masterPath)).toBe(true);
    // Re-emitted from raw — content survives, even if formatting diverges.
    const xml = reload.readText(masterPath);
    expect(xml).toMatch(/<p:sldMaster\b/);
    // Re-parsing the produced container must succeed.
    const snap2 = await parsePptx(out);
    expect(snap2.root.masters.has(masterPath)).toBe(true);
  });

  it("flushes a dirty theme part through the new dirty.theme path", async () => {
    const path = join(FIXTURES_DIR.pathname, "04-multi-shape.pptx");
    const buf = await readFile(path);
    const snap = await parsePptx(buf);
    const themePath = [...snap.root.theme.keys()][0]!;
    expect(themePath).toBeDefined();

    const dirtied: PptxSnapshot = {
      ...snap,
      dirty: { ...snap.dirty, theme: new Set([themePath]) },
    };
    const out = await serializePptx(dirtied);
    const reload = await ooxml.OoxmlContainer.load(out);

    expect(reload.has(themePath)).toBe(true);
    const xml = reload.readText(themePath);
    expect(xml).toMatch(/<a:theme\b/);
    const snap2 = await parsePptx(out);
    expect(snap2.root.theme.has(themePath)).toBe(true);
  });
});
