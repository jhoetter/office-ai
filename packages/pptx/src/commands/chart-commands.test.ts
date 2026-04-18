import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { PptxAgent } from "../agent/agent.js";
import type { ChartShape, TextShape } from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadChartAgent(): Promise<{ agent: PptxAgent; chart: ChartShape }> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, "09-with-chart.pptx"));
  const agent = await PptxAgent.fromBuffer(buf);
  const slide = agent.getSnapshot().root.slides[0]!;
  const chart = slide.shapes.find((s): s is ChartShape => s.kind === "chart")!;
  expect(chart).toBeDefined();
  return { agent, chart };
}

function _getChart(agent: PptxAgent): ChartShape {
  const slide = agent.getSnapshot().root.slides[0]!;
  return slide.shapes.find((s): s is ChartShape => s.kind === "chart")!;
}

describe("F3: pptx:set-chart-title", () => {
  it("updates the chart title and marks chart + slide dirty", async () => {
    const { agent, chart } = await loadChartAgent();
    await agent.applyCommand({
      type: "pptx:set-chart-title",
      payload: { slideIndex: 0, shapeId: chart.id, title: "Q4 results" },
    });
    const snap = agent.getSnapshot();
    const part = snap.root.charts.get(chart.chartPartPath)!;
    expect(part.title).toBe("Q4 results");
    expect(snap.dirty.charts.size).toBe(1);
    expect(snap.dirty.slides.size).toBe(1);
  });

  it("removes the title when passed null", async () => {
    const { agent, chart } = await loadChartAgent();
    // First set then remove.
    await agent.applyCommand({
      type: "pptx:set-chart-title",
      payload: { slideIndex: 0, shapeId: chart.id, title: "X" },
    });
    await agent.applyCommand({
      type: "pptx:set-chart-title",
      payload: { slideIndex: 0, shapeId: chart.id, title: null },
    });
    const part = agent.getSnapshot().root.charts.get(chart.chartPartPath)!;
    expect(part.title).toBeUndefined();
  });

  it("rejects no-op when title is unchanged", async () => {
    const { agent, chart } = await loadChartAgent();
    const partBefore = agent.getSnapshot().root.charts.get(chart.chartPartPath)!;
    const m = await agent.applyCommand({
      type: "pptx:set-chart-title",
      payload: { slideIndex: 0, shapeId: chart.id, title: partBefore.title ?? null },
    });
    expect(m.rejection?.code).toBe("no-op");
  });

  it("rejects non-chart shapes with not-applicable", async () => {
    const { agent } = await loadChartAgent();
    const slide = agent.getSnapshot().root.slides[0]!;
    const text = slide.shapes.find((s): s is TextShape => s.kind === "text");
    if (!text) return;
    const m = await agent.applyCommand({
      type: "pptx:set-chart-title",
      payload: { slideIndex: 0, shapeId: text.id, title: "x" },
    });
    expect(m.rejection?.code).toBe("not-applicable");
  });
});

describe("F3: pptx:set-chart-data", () => {
  it("replaces categories + series and dirties the chart", async () => {
    const { agent, chart } = await loadChartAgent();
    await agent.applyCommand({
      type: "pptx:set-chart-data",
      payload: {
        slideIndex: 0,
        shapeId: chart.id,
        categories: ["Jan", "Feb", "Mar"],
        series: [
          { name: "Foo", values: [1, 2, 3] },
          { name: "Bar", values: [4, 5, 6] },
        ],
      },
    });
    const part = agent.getSnapshot().root.charts.get(chart.chartPartPath)!;
    expect(part.categories).toEqual(["Jan", "Feb", "Mar"]);
    expect(part.series.length).toBe(2);
    expect(part.series[0]!.values).toEqual([1, 2, 3]);
    expect(part.series[1]!.name).toBe("Bar");
    expect(agent.getSnapshot().dirty.charts.has(chart.chartPartPath)).toBe(true);
  });

  it("rejects mismatched series length", async () => {
    const { agent, chart } = await loadChartAgent();
    const m = await agent.applyCommand({
      type: "pptx:set-chart-data",
      payload: {
        slideIndex: 0,
        shapeId: chart.id,
        categories: ["a", "b"],
        series: [{ values: [1, 2, 3] }],
      },
    });
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("rejects empty series list", async () => {
    const { agent, chart } = await loadChartAgent();
    const m = await agent.applyCommand({
      type: "pptx:set-chart-data",
      payload: { slideIndex: 0, shapeId: chart.id, categories: [], series: [] },
    });
    expect(m.rejection?.code).toBe("invalid-payload");
  });
});

describe("F3: pptx:set-chart-type", () => {
  it("changes chart type and dirties the chart", async () => {
    const { agent, chart } = await loadChartAgent();
    const before = agent.getSnapshot().root.charts.get(chart.chartPartPath)!.chartType;
    const target = before === "line" ? "bar" : "line";
    await agent.applyCommand({
      type: "pptx:set-chart-type",
      payload: { slideIndex: 0, shapeId: chart.id, chartType: target },
    });
    const part = agent.getSnapshot().root.charts.get(chart.chartPartPath)!;
    expect(part.chartType).toBe(target);
  });

  it("rejects no-op when chart already matches the requested type", async () => {
    const { agent, chart } = await loadChartAgent();
    const before = agent.getSnapshot().root.charts.get(chart.chartPartPath)!.chartType;
    if (before === "unsupported") return; // skip — meaningless for unsupported
    const m = await agent.applyCommand({
      type: "pptx:set-chart-type",
      payload: { slideIndex: 0, shapeId: chart.id, chartType: before as "bar" },
    });
    expect(m.rejection?.code).toBe("no-op");
  });

  it("rejects unsupported chart types", async () => {
    const { agent, chart } = await loadChartAgent();
    const m = await agent.applyCommand({
      type: "pptx:set-chart-type",
      payload: {
        slideIndex: 0,
        shapeId: chart.id,
        // @ts-expect-error testing runtime guard against unsupported
        chartType: "scatter",
      },
    });
    expect(m.rejection?.code).toBe("invalid-payload");
  });
});
