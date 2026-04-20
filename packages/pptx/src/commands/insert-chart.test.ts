import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { PptxAgent } from "../agent/agent.js";
import type { ChartShape } from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

describe("pptx:insert-chart", () => {
  it("authors a new ChartShape on the slide and registers a chart part", async () => {
    const buf = await readFile(join(FIXTURES_DIR.pathname, "04-multi-shape.pptx"));
    const agent = await PptxAgent.fromBuffer(buf);
    const before = agent.getSnapshot();
    const chartCountBefore = before.root.charts.size;

    const slide = before.root.slides[0]!;
    const shapesBefore = slide.shapes.length;

    await agent.applyCommand({
      type: "pptx:insert-chart",
      payload: {
        slideIndex: 0,
        x: 1_000_000,
        y: 1_000_000,
        chartType: "bar",
        title: "Quarterly revenue",
        categories: ["Q1", "Q2", "Q3", "Q4"],
        series: [
          { name: "EU", values: [10, 20, 30, 40] },
          { name: "US", values: [15, 25, 35, 45] },
        ],
      },
    });

    const after = agent.getSnapshot();
    const slideAfter = after.root.slides[0]!;
    expect(slideAfter.shapes.length).toBe(shapesBefore + 1);
    const chart = slideAfter.shapes[slideAfter.shapes.length - 1] as ChartShape;
    expect(chart.kind).toBe("chart");
    expect(chart.chartPartPath).toMatch(/^ppt\/charts\/chart\d+\.xml$/);
    expect(chart.chartRelId).toMatch(/^rId\d+$/);

    const part = after.root.charts.get(chart.chartPartPath)!;
    expect(part.chartType).toBe("bar");
    expect(part.title).toBe("Quarterly revenue");
    expect(part.categories).toEqual(["Q1", "Q2", "Q3", "Q4"]);
    expect(part.series.length).toBe(2);
    expect(part.series[0].values).toEqual([10, 20, 30, 40]);

    expect(after.root.charts.size).toBe(chartCountBefore + 1);
    expect(after.dirty.charts.has(chart.chartPartPath)).toBe(true);
    expect(after.dirty.slides.has(slide.partPath)).toBe(true);
    expect(after.dirty.contentTypes).toBe(true);
  });

  it("survives a serialize → reparse roundtrip with the embedded xlsx", async () => {
    const buf = await readFile(join(FIXTURES_DIR.pathname, "04-multi-shape.pptx"));
    const agent = await PptxAgent.fromBuffer(buf);
    await agent.applyCommand({
      type: "pptx:insert-chart",
      payload: {
        slideIndex: 0,
        x: 500_000,
        y: 500_000,
        chartType: "line",
        categories: ["Mon", "Tue", "Wed"],
        series: [{ name: "Visits", values: [3, 7, 11] }],
      },
    });
    const out = await agent.exportFile();
    const reagent = await PptxAgent.fromBuffer(new Uint8Array(out));
    const slide = reagent.getSnapshot().root.slides[0]!;
    const chart = slide.shapes.find((s): s is ChartShape => s.kind === "chart");
    expect(chart).toBeDefined();
    const part = reagent.getSnapshot().root.charts.get(chart!.chartPartPath)!;
    expect(part.chartType).toBe("line");
    expect(part.categories).toEqual(["Mon", "Tue", "Wed"]);
    expect(part.series[0]?.values).toEqual([3, 7, 11]);
    // Embedded xlsx must be wired so Office can "Edit Data" round-trip.
    expect(part.embeddingPartPath).toMatch(/Microsoft_Excel_Worksheet\d+\.xlsx$/);
  });

  it("rejects malformed payloads", async () => {
    const buf = await readFile(join(FIXTURES_DIR.pathname, "04-multi-shape.pptx"));
    const agent = await PptxAgent.fromBuffer(buf);
    const result = await agent.applyCommand({
      type: "pptx:insert-chart",
      payload: {
        slideIndex: 0,
        x: 0,
        y: 0,
        chartType: "bar",
        categories: ["A", "B"],
        series: [{ values: [1] }],
      },
    });
    expect(result.status).toBe("rejected");
    if (result.status === "rejected") {
      expect(result.rejection.message).toMatch(/values length/);
    }
  });
});
