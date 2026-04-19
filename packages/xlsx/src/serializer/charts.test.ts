import { readFile } from "node:fs/promises";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { CommandBus, type SerializedCommand } from "@officeai/core";
import { parseXlsx } from "../parser/parse.js";
import { allXlsxHandlers } from "../commands/registry.js";
import { serializeXlsx } from "./serialize.js";
import {
  CHART_CONTENT_TYPE,
  CHART_REL_TYPE,
  serializeChartPart,
} from "./charts.js";
import type { ChartKind, SheetChart, XlsxSnapshot } from "../model/types.js";

const here = dirname(fileURLToPath(import.meta.url));
const fixtures = resolve(here, "../../../../fixtures/xlsx/synthetic");

async function loadFixture(name: string): Promise<XlsxSnapshot> {
  const buf = new Uint8Array(await readFile(resolve(fixtures, name)));
  return parseXlsx(buf);
}

async function applyAddChart(
  snap: XlsxSnapshot,
  payload: { sheet: string; kind: ChartKind; dataRange: string; title?: string }
): Promise<XlsxSnapshot> {
  const bus = new CommandBus<XlsxSnapshot>(snap);
  bus.registerAll(allXlsxHandlers);
  const cmd: SerializedCommand = {
    type: "xlsx:add-chart",
    payload: { ...payload, hasHeaderRow: true, hasCategoryColumn: true },
  };
  const m = await bus.dispatch(cmd);
  if (m.status === "rejected") {
    throw new Error(`add-chart was rejected: ${m.rejection?.code} ${m.rejection?.message}`);
  }
  return m.after;
}

describe("serializeChartPart — XML shape", () => {
  it("emits c:barChart with col direction for a column chart", () => {
    const chart: SheetChart = {
      id: "chart-1",
      kind: "column",
      dataRange: "Sheet1!A1:C5",
      hasHeaderRow: true,
      hasCategoryColumn: true,
      title: "Sales by quarter",
      anchor: {
        fromRow: 0,
        fromCol: 4,
        fromOffsetXPx: 0,
        fromOffsetYPx: 0,
        widthPx: 480,
        heightPx: 280,
        editAs: "oneCell",
      },
    };
    const xml = serializeChartPart(chart);
    expect(xml).toMatch(/<c:barChart>/);
    expect(xml).toMatch(/<c:barDir val="col"\/>/);
    expect(xml).toMatch(/<c:title>/);
    expect(xml).toMatch(/Sales by quarter/);
    // Two data series (cols B and C), each pulling values from the
    // body rows of its column.
    expect(xml).toMatch(/Sheet1!B2:B5/);
    expect(xml).toMatch(/Sheet1!C2:C5/);
    // Series titles come from the header row.
    expect(xml).toMatch(/Sheet1!B1/);
    expect(xml).toMatch(/Sheet1!C1/);
    // Categories come from the leading column.
    expect(xml).toMatch(/Sheet1!A2:A5/);
  });

  it("emits c:lineChart for kind: line", () => {
    const chart: SheetChart = {
      id: "chart-2",
      kind: "line",
      dataRange: "Sheet1!A1:B4",
      hasHeaderRow: true,
      hasCategoryColumn: true,
      anchor: {
        fromRow: 0,
        fromCol: 0,
        fromOffsetXPx: 0,
        fromOffsetYPx: 0,
        widthPx: 200,
        heightPx: 200,
        editAs: "oneCell",
      },
    };
    const xml = serializeChartPart(chart);
    expect(xml).toMatch(/<c:lineChart>/);
    expect(xml).not.toMatch(/<c:barChart>/);
  });

  it("emits c:pieChart for kind: pie and omits category axes", () => {
    const chart: SheetChart = {
      id: "chart-3",
      kind: "pie",
      dataRange: "Sheet1!A1:B5",
      hasHeaderRow: true,
      hasCategoryColumn: true,
      anchor: {
        fromRow: 0,
        fromCol: 0,
        fromOffsetXPx: 0,
        fromOffsetYPx: 0,
        widthPx: 200,
        heightPx: 200,
        editAs: "oneCell",
      },
    };
    const xml = serializeChartPart(chart);
    expect(xml).toMatch(/<c:pieChart>/);
    expect(xml).not.toMatch(/<c:catAx>/);
    expect(xml).not.toMatch(/<c:valAx>/);
  });

  it("quotes sheet names that contain spaces", () => {
    const chart: SheetChart = {
      id: "chart-4",
      kind: "column",
      dataRange: "Has Space!A1:B3",
      hasHeaderRow: true,
      hasCategoryColumn: true,
      anchor: {
        fromRow: 0,
        fromCol: 0,
        fromOffsetXPx: 0,
        fromOffsetYPx: 0,
        widthPx: 200,
        heightPx: 200,
        editAs: "oneCell",
      },
    };
    const xml = serializeChartPart(chart);
    expect(xml).toMatch(/'Has Space'!B2:B3/);
  });

  it("escapes XML-significant characters in the title", () => {
    const chart: SheetChart = {
      id: "chart-5",
      kind: "column",
      dataRange: "Sheet1!A1:B3",
      hasHeaderRow: true,
      hasCategoryColumn: true,
      title: "A & B <vs> C",
      anchor: {
        fromRow: 0,
        fromCol: 0,
        fromOffsetXPx: 0,
        fromOffsetYPx: 0,
        widthPx: 200,
        heightPx: 200,
        editAs: "oneCell",
      },
    };
    const xml = serializeChartPart(chart);
    expect(xml).toMatch(/A &amp; B &lt;vs&gt; C/);
  });
});

describe("serializeXlsx — chart round-trip", () => {
  it("authoring an add-chart command writes xl/charts/chartN.xml + drawing graphicFrame + content-type", async () => {
    const snap = await loadFixture("01-single-sheet-numbers.xlsx");
    const sheetName = snap.root.sheets[0]!.name;

    const next = await applyAddChart(snap, {
      sheet: sheetName,
      kind: "column",
      dataRange: "A1:C5",
      title: "Round-trip me",
    });
    expect(next.root.sheets[0]!.charts.length).toBe(1);
    expect(next.dirty.drawings.has(snap.root.sheets[0]!.partPath)).toBe(true);

    const buf = await serializeXlsx(next);
    const reparsed = await parseXlsx(new Uint8Array(buf));

    // The chart payload should appear as an opaque part — we don't
    // currently re-hydrate chart definitions on parse, so this is
    // the right place to verify durability.
    const chartPaths = [...reparsed.container.parts.keys()].filter((p) =>
      /^xl\/charts\/chart\d+\.xml$/.test(p)
    );
    expect(chartPaths.length).toBeGreaterThanOrEqual(1);
    const chartXml = reparsed.container.readText(chartPaths[0]!);
    expect(chartXml).toMatch(/<c:barChart>/);
    expect(chartXml).toMatch(/Round-trip me/);
    expect(chartXml).toMatch(new RegExp(`${sheetName}!B2:B5`));

    // The drawing part should have been written too, with a
    // graphicFrame anchor pointing at our chart via a chart rel.
    const drawingPaths = [...reparsed.container.parts.keys()].filter((p) =>
      /^xl\/drawings\/drawing\d+\.xml$/.test(p)
    );
    expect(drawingPaths.length).toBeGreaterThanOrEqual(1);
    const drawingXml = reparsed.container.readText(drawingPaths[0]!);
    expect(drawingXml).toMatch(/<xdr:graphicFrame/);
    expect(drawingXml).toMatch(/<c:chart\b/);

    // Content-types must register the chart part — without this
    // override Excel/LibreOffice silently drop the chart on open.
    const contentTypes = reparsed.container.readText("[Content_Types].xml");
    expect(contentTypes).toContain(CHART_CONTENT_TYPE);

    // Drawing rels must include the chart rel type so the
    // r:id we emitted in the graphic frame resolves.
    const drawingRelsPaths = [...reparsed.container.parts.keys()].filter((p) =>
      /^xl\/drawings\/_rels\/drawing\d+\.xml\.rels$/.test(p)
    );
    expect(drawingRelsPaths.length).toBeGreaterThanOrEqual(1);
    const relsXml = reparsed.container.readText(drawingRelsPaths[0]!);
    expect(relsXml).toContain(CHART_REL_TYPE);
  });

  it("authoring then removing a chart leaves no orphan chart parts", async () => {
    const snap = await loadFixture("01-single-sheet-numbers.xlsx");
    const sheetName = snap.root.sheets[0]!.name;

    const withChart = await applyAddChart(snap, {
      sheet: sheetName,
      kind: "line",
      dataRange: "A1:B4",
      title: "Will go away",
    });

    const bus = new CommandBus<XlsxSnapshot>(withChart);
    bus.registerAll(allXlsxHandlers);
    const chartId = withChart.root.sheets[0]!.charts[0]!.id;
    const m = await bus.dispatch({
      type: "xlsx:remove-chart",
      payload: { sheet: sheetName, chartId },
    });

    const buf = await serializeXlsx(m.after);
    const reparsed = await parseXlsx(new Uint8Array(buf));

    const chartPaths = [...reparsed.container.parts.keys()].filter((p) =>
      /^xl\/charts\/chart\d+\.xml$/.test(p)
    );
    expect(chartPaths).toEqual([]);

    const contentTypes = reparsed.container.readText("[Content_Types].xml");
    expect(contentTypes).not.toContain(CHART_CONTENT_TYPE);
  });

  it("emits all four chart kinds successfully", async () => {
    const snap = await loadFixture("01-single-sheet-numbers.xlsx");
    const sheetName = snap.root.sheets[0]!.name;
    let cur = snap;
    for (const kind of ["column", "bar", "line", "pie"] as const) {
      cur = await applyAddChart(cur, {
        sheet: sheetName,
        kind,
        dataRange: "A1:B4",
        title: `${kind} chart`,
      });
    }
    const buf = await serializeXlsx(cur);
    const reparsed = await parseXlsx(new Uint8Array(buf));
    const chartPaths = [...reparsed.container.parts.keys()].filter((p) =>
      /^xl\/charts\/chart\d+\.xml$/.test(p)
    );
    expect(chartPaths.length).toBe(4);
    const xmls = chartPaths.map((p) => reparsed.container.readText(p));
    expect(xmls.some((x) => /<c:barChart>[\s\S]*<c:barDir val="col"\/>/.test(x))).toBe(true);
    expect(xmls.some((x) => /<c:barChart>[\s\S]*<c:barDir val="bar"\/>/.test(x))).toBe(true);
    expect(xmls.some((x) => /<c:lineChart>/.test(x))).toBe(true);
    expect(xmls.some((x) => /<c:pieChart>/.test(x))).toBe(true);
  });
});
