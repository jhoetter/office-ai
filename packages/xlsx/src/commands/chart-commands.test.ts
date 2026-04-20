import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { CommandBus, defaultIdMinter } from "@officeai/core";
import type { XlsxSnapshot } from "../model/types.js";
import { parseXlsx } from "../parser/index.js";
import { allXlsxHandlers } from "./registry.js";

const here = dirname(fileURLToPath(import.meta.url));
const fixtures = resolve(here, "../../../../fixtures/xlsx/synthetic");

async function loadFixture(name: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(resolve(fixtures, name)));
}

async function makeBus(fixture: string): Promise<{
  bus: CommandBus<XlsxSnapshot>;
  initial: XlsxSnapshot;
}> {
  const buf = await loadFixture(fixture);
  const initial = await parseXlsx(buf, { idMinter: defaultIdMinter });
  const bus = new CommandBus<XlsxSnapshot>(initial);
  bus.registerAll(allXlsxHandlers);
  return { bus, initial };
}

/**
 * Bootstrap a workbook that already contains one column chart so the
 * update tests don't need to reach back into `xlsx:add-chart` for
 * every case.
 */
async function withSeedChart(): Promise<{
  bus: CommandBus<XlsxSnapshot>;
  sheetName: string;
  chartId: string;
}> {
  const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
  const sheetName = initial.root.sheets[0]!.name;
  const add = await bus.dispatch({
    type: "xlsx:add-chart",
    payload: {
      sheet: sheetName,
      kind: "column",
      dataRange: "A1:B5",
      hasHeaderRow: true,
      hasCategoryColumn: true,
      title: "Original",
    },
  });
  expect(add.status).toBe("approved");
  const chartId = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.charts[0]!.id;
  return { bus, sheetName, chartId };
}

describe("xlsx:update-chart — happy path", () => {
  it("patches kind / dataRange / toggles / title in place", async () => {
    const { bus, sheetName, chartId } = await withSeedChart();

    const upd = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: {
        sheet: sheetName,
        chartId,
        kind: "line",
        dataRange: "C2:D8",
        hasHeaderRow: false,
        hasCategoryColumn: false,
        title: "Q4 trend",
      },
    });
    expect(upd.status).toBe("approved");

    const chart = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.charts[0]!;
    expect(chart.id).toBe(chartId);
    expect(chart.kind).toBe("line");
    expect(chart.dataRange).toBe("C2:D8");
    expect(chart.hasHeaderRow).toBe(false);
    expect(chart.hasCategoryColumn).toBe(false);
    expect(chart.title).toBe("Q4 trend");
  });

  it("preserves the anchor (move/resize stay separate concerns)", async () => {
    const { bus, sheetName, chartId } = await withSeedChart();
    const before = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.charts[0]!.anchor;

    const upd = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: { sheet: sheetName, chartId, kind: "bar" },
    });
    expect(upd.status).toBe("approved");

    const after = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.charts[0]!.anchor;
    expect(after).toEqual(before);
  });

  it("only writes fields that were provided", async () => {
    const { bus, sheetName, chartId } = await withSeedChart();

    const upd = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: { sheet: sheetName, chartId, kind: "pie" },
    });
    expect(upd.status).toBe("approved");

    const chart = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.charts[0]!;
    expect(chart.kind).toBe("pie");
    expect(chart.dataRange).toBe("A1:B5");
    expect(chart.title).toBe("Original");
    expect(chart.hasHeaderRow).toBe(true);
    expect(chart.hasCategoryColumn).toBe(true);
  });
});

describe("xlsx:update-chart — title clearing", () => {
  it("removes the title when payload.title is null", async () => {
    const { bus, sheetName, chartId } = await withSeedChart();

    const upd = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: { sheet: sheetName, chartId, title: null },
    });
    expect(upd.status).toBe("approved");

    const chart = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.charts[0]!;
    expect(chart.title).toBeUndefined();
  });

  it("removes the title when payload.title is the empty string", async () => {
    const { bus, sheetName, chartId } = await withSeedChart();

    const upd = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: { sheet: sheetName, chartId, title: "" },
    });
    expect(upd.status).toBe("approved");

    const chart = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.charts[0]!;
    expect(chart.title).toBeUndefined();
  });

  it("leaves an existing title intact when title is omitted", async () => {
    const { bus, sheetName, chartId } = await withSeedChart();

    const upd = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: { sheet: sheetName, chartId, kind: "bar" },
    });
    expect(upd.status).toBe("approved");

    const chart = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.charts[0]!;
    expect(chart.title).toBe("Original");
  });
});

describe("xlsx:update-chart — rejection paths", () => {
  it("rejects unknown chart ids", async () => {
    const { bus, sheetName } = await withSeedChart();

    const upd = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: { sheet: sheetName, chartId: "node-does-not-exist", kind: "bar" },
    });
    expect(upd.status).toBe("rejected");
    expect(upd.rejection?.code).toBe("unknown-chart");
  });

  it("rejects single-cell dataRange updates", async () => {
    const { bus, sheetName, chartId } = await withSeedChart();

    const upd = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: { sheet: sheetName, chartId, dataRange: "C3:C3" },
    });
    expect(upd.status).toBe("rejected");
    expect(upd.rejection?.code).toBe("invalid-range");
  });

  it("rejects a syntactically broken dataRange", async () => {
    const { bus, sheetName, chartId } = await withSeedChart();

    const upd = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: { sheet: sheetName, chartId, dataRange: "garbage" },
    });
    expect(upd.status).toBe("rejected");
    expect(upd.rejection?.code).toBe("invalid-range");
  });

  it("rejects payloads that don't actually specify a change", async () => {
    const { bus, sheetName, chartId } = await withSeedChart();

    const upd = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: { sheet: sheetName, chartId },
    });
    expect(upd.status).toBe("rejected");
    expect(upd.rejection?.code).toBe("no-op");
  });

  it("rejects updates that re-state the existing values verbatim", async () => {
    const { bus, sheetName, chartId } = await withSeedChart();
    const chart = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.charts[0]!;

    const upd = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: {
        sheet: sheetName,
        chartId,
        kind: chart.kind,
        dataRange: chart.dataRange,
        hasHeaderRow: chart.hasHeaderRow,
        hasCategoryColumn: chart.hasCategoryColumn,
        title: chart.title ?? null,
      },
    });
    expect(upd.status).toBe("rejected");
    expect(upd.rejection?.code).toBe("no-op");
  });
});

describe("xlsx:add-chart — style fields", () => {
  it("persists palette / legend / data-label / gridline / axis title flags when provided", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheetName = initial.root.sheets[0]!.name;

    const add = await bus.dispatch({
      type: "xlsx:add-chart",
      payload: {
        sheet: sheetName,
        kind: "column",
        dataRange: "A1:B5",
        hasHeaderRow: true,
        hasCategoryColumn: true,
        palette: "vibrant",
        showLegend: false,
        showDataLabels: true,
        showGridlines: false,
        xAxisTitle: "Quarter",
        yAxisTitle: "Revenue",
      },
    });
    expect(add.status).toBe("approved");

    const chart = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.charts[0]!;
    expect(chart.palette).toBe("vibrant");
    expect(chart.showLegend).toBe(false);
    expect(chart.showDataLabels).toBe(true);
    expect(chart.showGridlines).toBe(false);
    expect(chart.xAxisTitle).toBe("Quarter");
    expect(chart.yAxisTitle).toBe("Revenue");
  });

  it("omits style fields entirely when the payload doesn't mention them", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheetName = initial.root.sheets[0]!.name;

    const add = await bus.dispatch({
      type: "xlsx:add-chart",
      payload: {
        sheet: sheetName,
        kind: "column",
        dataRange: "A1:B5",
      },
    });
    expect(add.status).toBe("approved");

    const chart = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.charts[0]!;
    expect(chart.palette).toBeUndefined();
    expect(chart.showLegend).toBeUndefined();
    expect(chart.showDataLabels).toBeUndefined();
    expect(chart.showGridlines).toBeUndefined();
    expect(chart.xAxisTitle).toBeUndefined();
    expect(chart.yAxisTitle).toBeUndefined();
  });
});

describe("xlsx:update-chart — style fields", () => {
  it("patches palette + legend + data-labels + gridlines independently", async () => {
    const { bus, sheetName, chartId } = await withSeedChart();

    const upd = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: {
        sheet: sheetName,
        chartId,
        palette: "warm",
        showLegend: false,
        showDataLabels: true,
        showGridlines: false,
      },
    });
    expect(upd.status).toBe("approved");

    const chart = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.charts[0]!;
    expect(chart.palette).toBe("warm");
    expect(chart.showLegend).toBe(false);
    expect(chart.showDataLabels).toBe(true);
    expect(chart.showGridlines).toBe(false);
    // Existing typed fields untouched.
    expect(chart.kind).toBe("column");
    expect(chart.dataRange).toBe("A1:B5");
    expect(chart.title).toBe("Original");
  });

  it("clears palette/showLegend/etc back to undefined when given null", async () => {
    const { bus, sheetName, chartId } = await withSeedChart();

    const set = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: {
        sheet: sheetName,
        chartId,
        palette: "cool",
        showLegend: false,
        showDataLabels: true,
        showGridlines: false,
        xAxisTitle: "x",
        yAxisTitle: "y",
      },
    });
    expect(set.status).toBe("approved");

    const reset = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: {
        sheet: sheetName,
        chartId,
        palette: null,
        showLegend: null,
        showDataLabels: null,
        showGridlines: null,
        xAxisTitle: null,
        yAxisTitle: null,
      },
    });
    expect(reset.status).toBe("approved");

    const chart = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.charts[0]!;
    expect(chart.palette).toBeUndefined();
    expect(chart.showLegend).toBeUndefined();
    expect(chart.showDataLabels).toBeUndefined();
    expect(chart.showGridlines).toBeUndefined();
    expect(chart.xAxisTitle).toBeUndefined();
    expect(chart.yAxisTitle).toBeUndefined();
  });

  it("clears axis titles when given the empty string (parity with `title`)", async () => {
    const { bus, sheetName, chartId } = await withSeedChart();

    const set = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: { sheet: sheetName, chartId, xAxisTitle: "Quarter", yAxisTitle: "Revenue" },
    });
    expect(set.status).toBe("approved");

    const cleared = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: { sheet: sheetName, chartId, xAxisTitle: "", yAxisTitle: "" },
    });
    expect(cleared.status).toBe("approved");

    const chart = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.charts[0]!;
    expect(chart.xAxisTitle).toBeUndefined();
    expect(chart.yAxisTitle).toBeUndefined();
  });

  it("rejects no-op style updates that re-state existing values", async () => {
    const { bus, sheetName, chartId } = await withSeedChart();

    const set = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: { sheet: sheetName, chartId, palette: "mono", showLegend: false },
    });
    expect(set.status).toBe("approved");

    const noop = await bus.dispatch({
      type: "xlsx:update-chart",
      payload: { sheet: sheetName, chartId, palette: "mono", showLegend: false },
    });
    expect(noop.status).toBe("rejected");
    expect(noop.rejection?.code).toBe("no-op");
  });
});
