import { describe, expect, it } from "vitest";
import { CommandError, type CommandContext } from "@officeai/core";
import { addTableHandler } from "./add-table.js";
import { removeTableHandler } from "./remove-table.js";
import { cellKey } from "../model/refs.js";
import type { Cell, Sheet, XlsxSnapshot } from "../model/types.js";
import { defaultStyleTable } from "../model/style-table.js";

let nextId = 0;
const ctx: CommandContext = {
  mintNodeId: () => `node-${++nextId}`,
};

function makeSheet(name: string, cells: Iterable<[string, Cell]> = []): Sheet {
  return {
    id: `sheet-${name}`,
    sheetId: "1",
    name,
    index: 0,
    state: "visible",
    kind: "worksheet",
    partPath: `xl/worksheets/${name.toLowerCase()}.xml`,
    cells: new Map(cells),
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

function makeSnapshot(sheet: Sheet = makeSheet("Sheet1")): XlsxSnapshot {
  return {
    format: "xlsx",
    revision: 0,
    root: {
      id: "wb",
      sheets: [sheet],
      partHashes: {},
      opaqueParts: new Map(),
      date1904: false,
      workbookRootAttrs: {},
      styles: defaultStyleTable(),
      sheetjs: {
        SheetNames: [sheet.name],
        Sheets: { [sheet.name]: { "!ref": "A1" } },
      } as never,
      images: new Map(),
      definedNames: [],
    },
    partHashes: {},
    container: { parts: new Map() } as never,
    dirty: {
      workbook: false,
      sharedStrings: false,
      styles: false,
      contentTypes: false,
      rels: false,
      sheets: new Set(),
      comments: new Set(),
      threadedComments: new Set(),
      sheetRels: new Set(),
      removedSheetParts: new Set(),
      drawings: new Set(),
      media: new Set(),
      removedMediaParts: new Set(),
    },
  };
}

describe("xlsx:add-table", () => {
  it("adds a table with synthesised column names + AutoFilter", () => {
    const snap = makeSnapshot(
      makeSheet("Sheet1", [
        [cellKey(0, 0), { value: "Region", styleId: 0 }],
        [cellKey(0, 1), { value: "Sales", styleId: 0 }],
        [cellKey(1, 0), { value: "EU", styleId: 0 }],
        [cellKey(1, 1), { value: 100, styleId: 0 }],
      ])
    );
    const result = addTableHandler.apply(snap, { sheet: "Sheet1", range: "A1:B5" }, ctx);
    const sheet = result.next.root.sheets[0]!;
    expect(sheet.tables).toHaveLength(1);
    const t = sheet.tables[0]!;
    expect(t.displayName).toBe("Table1");
    expect(t.range).toBe("A1:B5");
    expect(t.columnNames).toEqual(["Region", "Sales"]);
    expect(t.headerRowCount).toBe(1);
    expect(sheet.autoFilter).toBeDefined();
    expect(sheet.autoFilter?.range).toEqual({ r1: 0, c1: 0, r2: 4, c2: 1 });
    expect(result.next.dirty.sheets.has(sheet.partPath)).toBe(true);
  });

  it("rejects single-row ranges when headers are present", () => {
    const snap = makeSnapshot();
    expect(() => addTableHandler.apply(snap, { sheet: "Sheet1", range: "A1:C1" }, ctx)).toThrow(CommandError);
  });

  it("rejects overlapping table ranges", () => {
    const snap = makeSnapshot();
    const first = addTableHandler.apply(snap, { sheet: "Sheet1", range: "A1:B5" }, ctx);
    expect(() => addTableHandler.apply(first.next, { sheet: "Sheet1", range: "B3:C8" }, ctx)).toThrow(
      /overlaps existing table/
    );
  });

  it("synthesises ColumnN names when headers are disabled", () => {
    const snap = makeSnapshot();
    const result = addTableHandler.apply(snap, { sheet: "Sheet1", range: "A1:C5", hasHeaders: false }, ctx);
    const t = result.next.root.sheets[0]!.tables[0]!;
    expect(t.columnNames).toEqual(["Column1", "Column2", "Column3"]);
    expect(t.headerRowCount).toBe(0);
  });
});

describe("xlsx:remove-table", () => {
  it("removes a table and clears its matching AutoFilter", () => {
    const snap = makeSnapshot();
    const added = addTableHandler.apply(snap, { sheet: "Sheet1", range: "A1:B5" }, ctx);
    const tableId = added.next.root.sheets[0]!.tables[0]!.tableId;
    const removed = removeTableHandler.apply(added.next, {
      sheet: "Sheet1",
      tableId,
    });
    expect(removed.next.root.sheets[0]!.tables).toHaveLength(0);
    expect(removed.next.root.sheets[0]!.autoFilter).toBeUndefined();
  });
});
