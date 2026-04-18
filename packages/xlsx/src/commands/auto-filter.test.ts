import { describe, expect, it } from "vitest";
import { CommandError } from "@officeai/core";
import { setAutoFilterHandler } from "./set-auto-filter.js";
import { setFilterColumnHandler } from "./set-filter-column.js";
import { clearFilterColumnHandler } from "./clear-filter-column.js";
import { sortRangeHandler } from "./sort-range.js";
import { recomputeHiddenRows } from "./auto-filter-eval.js";
import { cellKey } from "../model/refs.js";
import { defaultStyleTable } from "../model/style-table.js";
import type { Cell, Sheet, XlsxSnapshot } from "../model/types.js";

function makeSheet(name: string, rows: ReadonlyArray<ReadonlyArray<string | number | boolean | null>>): Sheet {
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
    images: [],
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

const sample = [
  ["Name", "Score", "Region"],
  ["Alice", 90, "North"],
  ["Bob", 70, "South"],
  ["Carol", 80, "North"],
  ["Dave", 60, "East"],
];

describe("xlsx:set-auto-filter", () => {
  it("installs an autoFilter band when given a range", () => {
    const snap = makeSnapshot(makeSheet("S", sample));
    const r = setAutoFilterHandler.apply(snap, { sheet: "S", range: "A1:C5" });
    const sheet = r.next.root.sheets[0]!;
    expect(sheet.autoFilter).toBeDefined();
    expect(sheet.autoFilter?.range).toEqual({ r1: 0, c1: 0, r2: 4, c2: 2 });
    expect(sheet.autoFilter?.columns.size).toBe(0);
    expect(sheet.hiddenRows.size).toBe(0);
  });

  it("rejects a range with no body row below the header", () => {
    const snap = makeSnapshot(makeSheet("S", sample));
    expect(() => setAutoFilterHandler.apply(snap, { sheet: "S", range: "A1:C1" })).toThrow(CommandError);
  });

  it("removes the autoFilter and clears hidden rows when range is null", () => {
    let snap = makeSnapshot(makeSheet("S", sample));
    snap = setAutoFilterHandler.apply(snap, { sheet: "S", range: "A1:C5" }).next;
    snap = setFilterColumnHandler.apply(snap, {
      sheet: "S",
      colId: 2,
      criterion: { kind: "values", values: new Set(["North"]), blank: false },
    }).next;
    expect(snap.root.sheets[0]!.hiddenRows.size).toBe(2);

    snap = setAutoFilterHandler.apply(snap, { sheet: "S", range: null }).next;
    expect(snap.root.sheets[0]!.autoFilter).toBeUndefined();
    expect(snap.root.sheets[0]!.hiddenRows.size).toBe(0);
  });
});

describe("xlsx:set-filter-column", () => {
  it("hides rows that don't match a values filter", () => {
    let snap = makeSnapshot(makeSheet("S", sample));
    snap = setAutoFilterHandler.apply(snap, { sheet: "S", range: "A1:C5" }).next;
    snap = setFilterColumnHandler.apply(snap, {
      sheet: "S",
      colId: 2,
      criterion: { kind: "values", values: new Set(["North"]), blank: false },
    }).next;
    expect([...snap.root.sheets[0]!.hiddenRows].sort()).toEqual([2, 4]);
  });

  it("respects custom > operator on numbers", () => {
    let snap = makeSnapshot(makeSheet("S", sample));
    snap = setAutoFilterHandler.apply(snap, { sheet: "S", range: "A1:C5" }).next;
    snap = setFilterColumnHandler.apply(snap, {
      sheet: "S",
      colId: 1,
      criterion: {
        kind: "custom",
        op1: { operator: "greaterThan", val: "75" },
        combine: "and",
      },
    }).next;
    expect([...snap.root.sheets[0]!.hiddenRows].sort()).toEqual([2, 4]);
  });

  it("supports top10 by items", () => {
    let snap = makeSnapshot(makeSheet("S", sample));
    snap = setAutoFilterHandler.apply(snap, { sheet: "S", range: "A1:C5" }).next;
    snap = setFilterColumnHandler.apply(snap, {
      sheet: "S",
      colId: 1,
      criterion: { kind: "top10", top: true, percent: false, n: 2, filterVal: 0 },
    }).next;
    expect([...snap.root.sheets[0]!.hiddenRows].sort()).toEqual([2, 4]);
  });

  it("rejects when no autoFilter is active", () => {
    const snap = makeSnapshot(makeSheet("S", sample));
    expect(() =>
      setFilterColumnHandler.apply(snap, {
        sheet: "S",
        colId: 0,
        criterion: { kind: "values", values: new Set(), blank: false },
      })
    ).toThrow(CommandError);
  });

  it("rejects out-of-range colId", () => {
    let snap = makeSnapshot(makeSheet("S", sample));
    snap = setAutoFilterHandler.apply(snap, { sheet: "S", range: "A1:C5" }).next;
    expect(() =>
      setFilterColumnHandler.apply(snap, {
        sheet: "S",
        colId: 99,
        criterion: { kind: "values", values: new Set(), blank: false },
      })
    ).toThrow(CommandError);
  });
});

describe("xlsx:clear-filter-column", () => {
  it("unhides rows previously hidden by that column", () => {
    let snap = makeSnapshot(makeSheet("S", sample));
    snap = setAutoFilterHandler.apply(snap, { sheet: "S", range: "A1:C5" }).next;
    snap = setFilterColumnHandler.apply(snap, {
      sheet: "S",
      colId: 2,
      criterion: { kind: "values", values: new Set(["North"]), blank: false },
    }).next;
    expect(snap.root.sheets[0]!.hiddenRows.size).toBeGreaterThan(0);
    snap = clearFilterColumnHandler.apply(snap, { sheet: "S", colId: 2 }).next;
    expect(snap.root.sheets[0]!.hiddenRows.size).toBe(0);
    expect(snap.root.sheets[0]!.autoFilter?.columns.has(2)).toBe(false);
  });
});

describe("xlsx:sort-range", () => {
  it("sorts body rows ascending by the chosen column", () => {
    const snap = makeSnapshot(makeSheet("S", sample));
    const r = sortRangeHandler.apply(snap, {
      sheet: "S",
      range: "A1:C5",
      sortBy: { colId: 1, order: "asc" },
    });
    const sheet = r.next.root.sheets[0]!;
    expect(sheet.cells.get(cellKey(1, 0))?.value).toBe("Dave");
    expect(sheet.cells.get(cellKey(2, 0))?.value).toBe("Bob");
    expect(sheet.cells.get(cellKey(3, 0))?.value).toBe("Carol");
    expect(sheet.cells.get(cellKey(4, 0))?.value).toBe("Alice");
  });

  it("sorts descending when order=desc and keeps header", () => {
    const snap = makeSnapshot(makeSheet("S", sample));
    const r = sortRangeHandler.apply(snap, {
      sheet: "S",
      range: "A1:C5",
      sortBy: { colId: 1, order: "desc" },
    });
    const sheet = r.next.root.sheets[0]!;
    expect(sheet.cells.get(cellKey(0, 0))?.value).toBe("Name");
    expect(sheet.cells.get(cellKey(1, 0))?.value).toBe("Alice");
    expect(sheet.cells.get(cellKey(4, 0))?.value).toBe("Dave");
  });

  it("returns a no-op when the range is already sorted", () => {
    const sortedRows = [
      ["Name", "Score"],
      ["A", 1],
      ["B", 2],
      ["C", 3],
    ];
    const snap = makeSnapshot(makeSheet("S", sortedRows));
    const r = sortRangeHandler.apply(snap, {
      sheet: "S",
      range: "A1:B4",
      sortBy: { colId: 1, order: "asc" },
    });
    expect(r.diff.changes).toHaveLength(0);
  });

  it("recomputes hidden rows when an autoFilter is active", () => {
    let snap = makeSnapshot(makeSheet("S", sample));
    snap = setAutoFilterHandler.apply(snap, { sheet: "S", range: "A1:C5" }).next;
    snap = setFilterColumnHandler.apply(snap, {
      sheet: "S",
      colId: 2,
      criterion: { kind: "values", values: new Set(["North"]), blank: false },
    }).next;
    snap = sortRangeHandler.apply(snap, {
      sheet: "S",
      range: "A1:C5",
      sortBy: { colId: 0, order: "asc" },
    }).next;
    const sheet = snap.root.sheets[0]!;
    const visible: string[] = [];
    for (let r = 1; r <= 4; r++) {
      if (!sheet.hiddenRows.has(r)) {
        visible.push(sheet.cells.get(cellKey(r, 2))?.value as string);
      }
    }
    expect(visible.every((v) => v === "North")).toBe(true);
    expect(visible.length).toBe(2);
  });
});

describe("recomputeHiddenRows", () => {
  it("returns an empty set when autoFilter is undefined", () => {
    const sheet = makeSheet("S", sample);
    const hidden = recomputeHiddenRows(sheet, defaultStyleTable(), undefined);
    expect(hidden.size).toBe(0);
  });

  it("ANDs criteria across columns", () => {
    let snap = makeSnapshot(makeSheet("S", sample));
    snap = setAutoFilterHandler.apply(snap, { sheet: "S", range: "A1:C5" }).next;
    snap = setFilterColumnHandler.apply(snap, {
      sheet: "S",
      colId: 1,
      criterion: {
        kind: "custom",
        op1: { operator: "greaterThanOrEqual", val: "70" },
        combine: "and",
      },
    }).next;
    snap = setFilterColumnHandler.apply(snap, {
      sheet: "S",
      colId: 2,
      criterion: { kind: "values", values: new Set(["North"]), blank: false },
    }).next;
    expect([...snap.root.sheets[0]!.hiddenRows].sort()).toEqual([2, 4]);
  });
});
