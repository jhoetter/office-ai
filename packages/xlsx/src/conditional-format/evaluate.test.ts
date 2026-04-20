import { describe, expect, it } from "vitest";
import type { Cell, Sheet } from "../model/types.js";
import { evaluateConditionalFormats } from "./evaluate.js";

function num(v: number): Cell {
  return { type: "number", value: v };
}
function str(v: string): Cell {
  return { type: "string", value: v };
}

function makeSheet(cells: ReadonlyMap<string, Cell>, rules: Sheet["conditionalFormats"]): Sheet {
  return {
    id: "s1",
    sheetId: "1",
    name: "Sheet1",
    index: 0,
    state: "visible",
    kind: "worksheet",
    partPath: "xl/worksheets/sheet1.xml",
    cells,
    merges: [],
    comments: [],
    commentAuthors: [],
    columnWidths: new Map(),
    rowHeights: new Map(),
    hiddenRows: new Set(),
    hiddenCols: new Set(),
    images: [],
    conditionalFormats: rules,
    opaqueConditionalFormats: [],
    dataValidations: [],
    tables: [],
    charts: [],
  };
}

describe("evaluateConditionalFormats", () => {
  it("applies cellIs > rules to numeric cells", () => {
    const cells = new Map<string, Cell>([
      ["0:0", num(10)],
      ["0:1", num(50)],
      ["0:2", num(99)],
    ]);
    const sheet = makeSheet(cells, [
      {
        kind: "cellIs",
        id: "r1",
        range: "A1:C1",
        op: "gt",
        value: 40,
        overlay: { fill: "FFFF00" },
      },
    ]);
    const overlays = evaluateConditionalFormats(sheet);
    expect(overlays.get("0:0")).toBeUndefined();
    expect(overlays.get("0:1")?.fill).toBe("FFFF00");
    expect(overlays.get("0:2")?.fill).toBe("FFFF00");
  });

  it("applies top10 ranking", () => {
    const cells = new Map<string, Cell>([
      ["0:0", num(10)],
      ["1:0", num(40)],
      ["2:0", num(80)],
      ["3:0", num(20)],
    ]);
    const sheet = makeSheet(cells, [
      {
        kind: "top10",
        id: "t1",
        range: "A1:A4",
        bottom: false,
        percent: false,
        rank: 2,
        overlay: { fontColor: "FF0000" },
      },
    ]);
    const o = evaluateConditionalFormats(sheet);
    expect(o.get("2:0")?.fontColor).toBe("FF0000");
    expect(o.get("1:0")?.fontColor).toBe("FF0000");
    expect(o.get("3:0")?.fontColor).toBeUndefined();
  });

  it("flags duplicates", () => {
    const cells = new Map<string, Cell>([
      ["0:0", str("apple")],
      ["1:0", str("banana")],
      ["2:0", str("apple")],
    ]);
    const sheet = makeSheet(cells, [
      {
        kind: "duplicate",
        id: "d1",
        range: "A1:A3",
        unique: false,
        overlay: { fill: "FF0000" },
      },
    ]);
    const o = evaluateConditionalFormats(sheet);
    expect(o.get("0:0")?.fill).toBe("FF0000");
    expect(o.get("2:0")?.fill).toBe("FF0000");
    expect(o.get("1:0")?.fill).toBeUndefined();
  });

  it("colors a 3-stop colour scale", () => {
    const cells = new Map<string, Cell>([
      ["0:0", num(0)],
      ["0:1", num(50)],
      ["0:2", num(100)],
    ]);
    const sheet = makeSheet(cells, [
      {
        kind: "colorScale",
        id: "cs",
        range: "A1:C1",
        minColor: "FF0000",
        midColor: "FFFF00",
        maxColor: "00FF00",
      },
    ]);
    const o = evaluateConditionalFormats(sheet);
    expect(o.get("0:0")?.fill).toBe("FF0000");
    expect(o.get("0:1")?.fill).toBe("FFFF00");
    expect(o.get("0:2")?.fill).toBe("00FF00");
  });

  it("computes data-bar fractions", () => {
    const cells = new Map<string, Cell>([
      ["0:0", num(0)],
      ["1:0", num(50)],
      ["2:0", num(100)],
    ]);
    const sheet = makeSheet(cells, [
      {
        kind: "dataBar",
        id: "db",
        range: "A1:A3",
        color: "638EC6",
      },
    ]);
    const o = evaluateConditionalFormats(sheet);
    expect(o.get("0:0")?.barFraction).toBeCloseTo(0, 5);
    expect(o.get("1:0")?.barFraction).toBeCloseTo(0.5, 5);
    expect(o.get("2:0")?.barFraction).toBeCloseTo(1, 5);
  });
});
