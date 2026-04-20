import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { parseXlsx } from "../parser/index.js";
import { cellKey } from "../model/refs.js";
import { buildChartGrid, buildEmbeddedXlsx } from "./buildEmbeddedXlsx.js";

describe("buildEmbeddedXlsx", () => {
  it("produces a valid xlsx package re-readable by parseXlsx", async () => {
    const grid = [
      ["", "Q1", "Q2"],
      ["Apples", 10, 20],
      ["Oranges", 15, 25],
      ["Pears", 12, 18],
    ];
    const result = await buildEmbeddedXlsx(grid);
    expect(result.bytes.byteLength).toBeGreaterThan(200);
    expect(result.sheetName).toBe("Sheet1");

    const z = await JSZip.loadAsync(result.bytes);
    expect(z.file("[Content_Types].xml")).toBeTruthy();
    expect(z.file("xl/workbook.xml")).toBeTruthy();
    expect(z.file("xl/worksheets/sheet1.xml")).toBeTruthy();

    const snap = await parseXlsx(result.bytes);
    const sheet = snap.root.sheets[0]!;
    expect(sheet.name).toBe("Sheet1");
    const b1 = sheet.cells.get(cellKey(0, 1));
    expect(b1?.value).toBe("Q1");
    const a2 = sheet.cells.get(cellKey(1, 0));
    expect(a2?.value).toBe("Apples");
    const b2 = sheet.cells.get(cellKey(1, 1));
    expect(b2?.value).toBe(10);
  });

  it("supports custom sheet name", async () => {
    const result = await buildEmbeddedXlsx([["x"]], { sheetName: "Daten" });
    expect(result.sheetName).toBe("Daten");
    const snap = await parseXlsx(result.bytes);
    expect(snap.root.sheets[0]?.name).toBe("Daten");
  });

  it("rejects sheet names with forbidden characters", async () => {
    await expect(buildEmbeddedXlsx([["x"]], { sheetName: "Bad/Name" })).rejects.toThrow();
  });
});

describe("buildChartGrid", () => {
  it("lays out series columns with correct cell refs", () => {
    const result = buildChartGrid(
      ["Apples", "Oranges", "Pears"],
      [
        { name: "Q1", values: [10, 15, 12] },
        { name: "Q2", values: [20, 25, 18] },
      ]
    );
    expect(result.grid).toEqual([
      ["", "Q1", "Q2"],
      ["Apples", 10, 20],
      ["Oranges", 15, 25],
      ["Pears", 12, 18],
    ]);
    expect(result.categoryRef).toBe("Sheet1!$A$2:$A$4");
    expect(result.valueRefs).toEqual(["Sheet1!$B$2:$B$4", "Sheet1!$C$2:$C$4"]);
    expect(result.nameRefs).toEqual(["Sheet1!$B$1", "Sheet1!$C$1"]);
  });

  it("synthesises a series name when missing", () => {
    const result = buildChartGrid(["a"], [{ values: [1] }]);
    expect(result.grid[0]).toEqual(["", "Series 1"]);
  });
});
