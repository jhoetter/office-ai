import { describe, expect, it } from "vitest";
import { CommandError } from "@officeai/core";
import { removeDuplicatesHandler } from "./remove-duplicates.js";
import { cellKey } from "../model/refs.js";
import { defaultStyleTable } from "../model/style-table.js";
import type { Cell, Sheet, XlsxSnapshot } from "../model/types.js";

function makeSheet(
  name: string,
  rows: ReadonlyArray<ReadonlyArray<string | number | boolean | null>>
): Sheet {
  const cells = new Map<string, Cell>();
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r]!;
    for (let c = 0; c < row.length; c++) {
      const v = row[c];
      if (v === null || v === undefined) continue;
      cells.set(cellKey(r, c), { row: r, col: c, value: v });
    }
  }
  return {
    id: "sheet-1",
    sheetId: "1",
    name,
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
    conditionalFormats: [],
    opaqueConditionalFormats: [],
    dataValidations: [],
    tables: [],
    charts: [],
  };
}

function makeSnapshot(sheet: Sheet): XlsxSnapshot {
  return {
    revision: 0,
    root: {
      id: "wb",
      sheets: [sheet],
      activeSheetId: sheet.id,
      definedNames: [],
      styles: defaultStyleTable(),
      sheetjs: { SheetNames: [sheet.name], Sheets: { [sheet.name]: { "!ref": "A1" } } },
    } as XlsxSnapshot["root"],
    dirty: {
      sheets: new Set(),
      workbook: false,
      sharedStrings: false,
      styles: false,
      contentTypes: false,
      rels: false,
    },
  } as unknown as XlsxSnapshot;
}

describe("xlsx:remove-duplicates", () => {
  it("removes duplicate rows comparing every column by default", () => {
    const snap = makeSnapshot(
      makeSheet("S", [
        ["Name", "Region"],
        ["Alice", "North"],
        ["Bob", "South"],
        ["Alice", "North"],
        ["Carol", "East"],
        ["Bob", "South"],
      ])
    );
    const out = removeDuplicatesHandler.apply(snap, { sheet: "S", range: "A1:B6" });
    const sheet = out.next.root.sheets[0]!;
    expect(sheet.cells.get(cellKey(1, 0))?.value).toBe("Alice");
    expect(sheet.cells.get(cellKey(2, 0))?.value).toBe("Bob");
    expect(sheet.cells.get(cellKey(3, 0))?.value).toBe("Carol");
    // trailing rows cleared
    expect(sheet.cells.get(cellKey(4, 0))).toBeUndefined();
    expect(sheet.cells.get(cellKey(5, 0))).toBeUndefined();
  });

  it("compares strings case-insensitively", () => {
    const snap = makeSnapshot(makeSheet("S", [["A"], ["alpha"], ["ALPHA"], ["beta"]]));
    const out = removeDuplicatesHandler.apply(snap, { sheet: "S", range: "A1:A4" });
    const sheet = out.next.root.sheets[0]!;
    expect(sheet.cells.get(cellKey(1, 0))?.value).toBe("alpha");
    expect(sheet.cells.get(cellKey(2, 0))?.value).toBe("beta");
    expect(sheet.cells.get(cellKey(3, 0))).toBeUndefined();
  });

  it("dedupes by selected key columns only", () => {
    // Same Region collapses; Score column is irrelevant.
    const snap = makeSnapshot(
      makeSheet("S", [
        ["Name", "Score", "Region"],
        ["Alice", 90, "North"],
        ["Bob", 70, "South"],
        ["Carol", 60, "North"],
        ["Dave", 50, "East"],
      ])
    );
    const out = removeDuplicatesHandler.apply(snap, {
      sheet: "S",
      range: "A1:C5",
      keyCols: [2],
    });
    const sheet = out.next.root.sheets[0]!;
    expect(sheet.cells.get(cellKey(1, 2))?.value).toBe("North");
    expect(sheet.cells.get(cellKey(2, 2))?.value).toBe("South");
    expect(sheet.cells.get(cellKey(3, 2))?.value).toBe("East");
    expect(sheet.cells.get(cellKey(4, 0))).toBeUndefined();
  });

  it("returns a no-op diff when nothing changes", () => {
    const snap = makeSnapshot(makeSheet("S", [["A"], ["x"], ["y"], ["z"]]));
    const out = removeDuplicatesHandler.apply(snap, { sheet: "S", range: "A1:A4" });
    const sheet = out.next.root.sheets[0]!;
    expect(sheet.cells.get(cellKey(1, 0))?.value).toBe("x");
    expect(sheet.cells.get(cellKey(3, 0))?.value).toBe("z");
  });

  it("rejects keyCols offsets outside the range", () => {
    const snap = makeSnapshot(
      makeSheet("S", [
        ["A", "B"],
        ["x", "y"],
      ])
    );
    expect(() => removeDuplicatesHandler.apply(snap, { sheet: "S", range: "A1:B2", keyCols: [5] })).toThrow(
      CommandError
    );
  });
});
