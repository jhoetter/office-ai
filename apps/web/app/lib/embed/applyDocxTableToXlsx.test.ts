import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { XlsxAgent } from "@officeai/xlsx";
import { cellKey, parseA1 } from "@officeai/xlsx";
import { applyDocxTableToXlsx } from "./applyDocxTableToXlsx";

function valueAt(agent: XlsxAgent, sheetName: string, ref: string): unknown {
  const snap = agent.getSnapshot();
  const sheet = snap.root.sheets.find((s) => s.name === sheetName);
  if (!sheet) return undefined;
  const { row, col } = parseA1(ref);
  return sheet.cells.get(cellKey(row, col))?.value;
}

function makeAgent(): Promise<XlsxAgent> {
  return XlsxAgent.empty({ idMinter: deterministicIdMinter("docx-table-") });
}

describe("applyDocxTableToXlsx", () => {
  it("writes a 2D string matrix as a real range anchored at target", async () => {
    const agent = await makeAgent();
    const sheet = agent.listSheets()[0]!.name;
    await applyDocxTableToXlsx({
      agent,
      sheet,
      target: "B3",
      cells: [
        ["Name", "Score"],
        ["Ada", "99"],
        ["Linus", "42"],
      ],
    });
    expect(valueAt(agent, sheet, "B3")).toBe("Name");
    expect(valueAt(agent, sheet, "C3")).toBe("Score");
    expect(valueAt(agent, sheet, "B4")).toBe("Ada");
    // Numeric strings get coerced to numbers, matching the in-app
    // HTML / TSV importer.
    expect(valueAt(agent, sheet, "C4")).toBe(99);
    expect(valueAt(agent, sheet, "B5")).toBe("Linus");
    expect(valueAt(agent, sheet, "C5")).toBe(42);
  });

  it("coerces TRUE / FALSE strings to booleans (case-insensitive)", async () => {
    const agent = await makeAgent();
    const sheet = agent.listSheets()[0]!.name;
    await applyDocxTableToXlsx({
      agent,
      sheet,
      target: "A1",
      cells: [
        ["TRUE", "false"],
        ["True", "FALSE"],
      ],
    });
    expect(valueAt(agent, sheet, "A1")).toBe(true);
    expect(valueAt(agent, sheet, "B1")).toBe(false);
    expect(valueAt(agent, sheet, "A2")).toBe(true);
    expect(valueAt(agent, sheet, "B2")).toBe(false);
  });

  it("treats empty strings as sparse cells, matching XLSX paste-range semantics", async () => {
    const agent = await makeAgent();
    const sheet = agent.listSheets()[0]!.name;
    // Seed C2 with a value that should NOT be touched: the docx
    // table only has 2 columns (A:B) so C is outside the paste rect
    // entirely, even with sparse mapping.
    await agent.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet, ref: "C2", value: "kept" },
      source: "human",
    });
    await applyDocxTableToXlsx({
      agent,
      sheet,
      target: "A1",
      cells: [
        ["x", ""],
        ["", "y"],
      ],
    });
    expect(valueAt(agent, sheet, "A1")).toBe("x");
    expect(valueAt(agent, sheet, "B2")).toBe("y");
    expect(valueAt(agent, sheet, "C2")).toBe("kept");
  });

  it("pads rows so the snapshot is rectangular", async () => {
    const agent = await makeAgent();
    const sheet = agent.listSheets()[0]!.name;
    await applyDocxTableToXlsx({
      agent,
      sheet,
      target: "A1",
      cells: [
        ["a"],
        ["b", "c", "d"],
        ["e", "f"],
      ],
    });
    expect(valueAt(agent, sheet, "A1")).toBe("a");
    expect(valueAt(agent, sheet, "C2")).toBe("d");
    expect(valueAt(agent, sheet, "A3")).toBe("e");
    expect(valueAt(agent, sheet, "B3")).toBe("f");
  });

  it("is a no-op for an empty matrix", async () => {
    const agent = await makeAgent();
    const sheet = agent.listSheets()[0]!.name;
    await applyDocxTableToXlsx({ agent, sheet, target: "A1", cells: [] });
    expect(valueAt(agent, sheet, "A1")).toBeUndefined();
  });
});
