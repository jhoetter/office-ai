import { describe, expect, it } from "vitest";
import { deterministicIdMinter, ooxml } from "@officeai/core";
import * as JSZip from "jszip";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import { serializeDocx } from "../serializer/serialize.js";
import type { BlockNode, ChartDrawing, Paragraph } from "../model/types.js";
import { DEFAULT_DOC_ROOT_ATTRS, escapeXml, makeSyntheticDocx } from "../test-utils/synthetic.js";

function plainDoc(paragraphs: ReadonlyArray<string>): string {
  const ps = paragraphs
    .map((t) => `<w:p><w:r><w:t xml:space="preserve">${escapeXml(t)}</w:t></w:r></w:p>`)
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>${ps}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
}

async function loadAgent(paragraphs: ReadonlyArray<string>): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: plainDoc(paragraphs) });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function findChartDrawing(p: Paragraph): ChartDrawing | null {
  for (const inline of p.children) {
    if (inline.kind !== "run") continue;
    for (const c of inline.children) {
      if (c.kind === "drawing" && c.subkind === "chart") return c;
    }
  }
  return null;
}

function paraOf(block: BlockNode): Paragraph {
  if (block.kind !== "paragraph") throw new Error("expected paragraph");
  return block;
}

describe("docx charts — insert-chart authoring + roundtrip", () => {
  it("inserts a typed chart, mints a chart part + relationship, and dirty-flags the right things", async () => {
    const agent = await loadAgent(["Sales overview"]);

    const m = await agent.applyCommand({
      type: "docx:insert-chart",
      payload: {
        at: { paragraph: 0 },
        chartType: "bar",
        title: "Q1 Revenue",
        categories: ["Jan", "Feb", "Mar"],
        series: [
          { name: "Revenue", values: [10, 20, 30] },
          { name: "Cost", values: [5, 8, 12] },
        ],
      },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const snap = agent.getSnapshot();
    expect(snap.dirty.body).toBe(true);
    expect(snap.dirty.charts.size).toBe(1);
    expect(snap.dirty.relationships.has("word/document.xml")).toBe(true);
    expect(snap.dirty.contentTypes).toBe(true);

    const partPath = [...snap.dirty.charts][0]!;
    expect(partPath).toBe("word/charts/chart1.xml");
    const chart = snap.root.charts.get(partPath);
    expect(chart).toBeTruthy();
    expect(chart?.chartType).toBe("bar");
    expect(chart?.title).toBe("Q1 Revenue");
    expect(chart?.categories).toEqual(["Jan", "Feb", "Mar"]);
    expect(chart?.series).toHaveLength(2);

    const drawing = findChartDrawing(paraOf(snap.root.body[0]));
    expect(drawing).toBeTruthy();
    expect(drawing?.chartPartPath).toBe(partPath);

    expect(m.diff.changes.map((c) => c.kind).sort()).toEqual(["node-inserted", "part-added"]);
  });

  it("rejects mismatched series.values length", async () => {
    const agent = await loadAgent(["x"]);
    const m = await agent.applyCommand({
      type: "docx:insert-chart",
      payload: {
        at: { paragraph: 0 },
        chartType: "bar",
        categories: ["A", "B", "C"],
        series: [{ name: "S", values: [1, 2] }],
      },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    if (m.status === "rejected") {
      expect(m.rejection.code).toBe("invalid-payload");
      expect(m.rejection.message).toMatch(/values length/);
    }
  });

  it("serializes a freshly-inserted chart with an embedded xlsx workbook and roundtrips back into the typed model", async () => {
    const agent = await loadAgent(["Sales"]);
    await agent.applyCommand({
      type: "docx:insert-chart",
      payload: {
        at: { paragraph: 0 },
        chartType: "line",
        title: "Trend",
        categories: ["Q1", "Q2", "Q3", "Q4"],
        series: [{ name: "Units", values: [100, 120, 90, 140] }],
      },
      source: "human",
    });

    const out = await serializeDocx(agent.getSnapshot());

    const reloaded = await ooxml.OoxmlContainer.load(out);

    expect(reloaded.has("word/charts/chart1.xml")).toBe(true);
    const chartXml = reloaded.readText("word/charts/chart1.xml");
    expect(chartXml).toContain("<c:lineChart");
    expect(chartXml).toContain("Trend");
    expect(chartXml).toContain("c:externalData");

    const chartRels = reloaded.has("word/charts/_rels/chart1.xml.rels");
    expect(chartRels).toBe(true);
    const relsXml = reloaded.readText("word/charts/_rels/chart1.xml.rels");
    expect(relsXml).toContain("Microsoft_Excel_Worksheet1.xlsx");
    expect(relsXml).toContain("/relationships/package");

    const embeddingPaths = [...reloaded.parts.keys()].filter((p) =>
      p.startsWith("word/embeddings/Microsoft_Excel_Worksheet")
    );
    expect(embeddingPaths).toHaveLength(1);
    const embeddedBytes = reloaded.readBytes(embeddingPaths[0]!);
    expect(embeddedBytes.byteLength).toBeGreaterThan(0);

    const inner = await JSZip.loadAsync(embeddedBytes);
    expect(inner.file("[Content_Types].xml")).toBeTruthy();
    expect(inner.file("xl/workbook.xml")).toBeTruthy();
    expect(inner.file("xl/worksheets/sheet1.xml")).toBeTruthy();

    const ctXml = reloaded.readText("[Content_Types].xml");
    expect(ctXml).toContain("application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
    expect(ctXml).toContain("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

    const docRels = reloaded.readText("word/_rels/document.xml.rels");
    expect(docRels).toContain("/relationships/chart");
    expect(docRels).toContain("charts/chart1.xml");

    const reparsed = await parseDocx(out, { idMinter: deterministicIdMinter() });
    const reparsedChart = reparsed.root.charts.get("word/charts/chart1.xml");
    expect(reparsedChart).toBeTruthy();
    expect(reparsedChart?.chartType).toBe("line");
    expect(reparsedChart?.title).toBe("Trend");
    expect(reparsedChart?.categories).toEqual(["Q1", "Q2", "Q3", "Q4"]);
    expect(reparsedChart?.series[0]?.name).toBe("Units");
    expect(reparsedChart?.series[0]?.values).toEqual([100, 120, 90, 140]);
    expect(reparsedChart?.embeddingPartPath).toBe(embeddingPaths[0]);
  });

  it("set-chart-data replaces categories + series and refreshes the embedded workbook", async () => {
    const agent = await loadAgent(["x"]);
    await agent.applyCommand({
      type: "docx:insert-chart",
      payload: {
        at: { paragraph: 0 },
        chartType: "bar",
        categories: ["A", "B"],
        series: [{ values: [1, 2] }],
      },
      source: "human",
    });

    const partPath = "word/charts/chart1.xml";
    const m = await agent.applyCommand({
      type: "docx:set-chart-data",
      payload: {
        chartPartPath: partPath,
        categories: ["A", "B", "C"],
        series: [
          { name: "X", values: [10, 20, 30] },
          { name: "Y", values: [4, 5, 6] },
        ],
      },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const snap = agent.getSnapshot();
    const part = snap.root.charts.get(partPath);
    expect(part?.categories).toEqual(["A", "B", "C"]);
    expect(part?.series).toHaveLength(2);
    expect(part?.series[1]?.values).toEqual([4, 5, 6]);
  });

  it("set-chart-type switches the active plot type", async () => {
    const agent = await loadAgent(["x"]);
    await agent.applyCommand({
      type: "docx:insert-chart",
      payload: {
        at: { paragraph: 0 },
        chartType: "bar",
        categories: ["A"],
        series: [{ values: [1] }],
      },
      source: "human",
    });
    const partPath = "word/charts/chart1.xml";
    await agent.applyCommand({
      type: "docx:set-chart-type",
      payload: { chartPartPath: partPath, chartType: "pie" },
      source: "human",
    });
    expect(agent.getSnapshot().root.charts.get(partPath)?.chartType).toBe("pie");
  });

  it("set-chart-title can set and clear the title", async () => {
    const agent = await loadAgent(["x"]);
    await agent.applyCommand({
      type: "docx:insert-chart",
      payload: {
        at: { paragraph: 0 },
        chartType: "bar",
        categories: ["A"],
        series: [{ values: [1] }],
      },
      source: "human",
    });
    const partPath = "word/charts/chart1.xml";
    await agent.applyCommand({
      type: "docx:set-chart-title",
      payload: { chartPartPath: partPath, title: "Hello" },
      source: "human",
    });
    expect(agent.getSnapshot().root.charts.get(partPath)?.title).toBe("Hello");
    await agent.applyCommand({
      type: "docx:set-chart-title",
      payload: { chartPartPath: partPath, title: null },
      source: "human",
    });
    expect(agent.getSnapshot().root.charts.get(partPath)?.title).toBeUndefined();
  });
});
