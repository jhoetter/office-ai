import { describe, expect, it } from "vitest";
import { CommandError } from "@officeai/core";
import { setColumnWidthHandler } from "./set-column-width.js";
import { setRowHeightHandler } from "./set-row-height.js";
import type { Sheet, XlsxSnapshot } from "../model/types.js";
import { defaultStyleTable } from "../model/style-table.js";

function makeSheet(name: string): Sheet {
  return {
    id: "sheet-1",
    sheetId: "1",
    name,
    index: 0,
    state: "visible",
    kind: "worksheet",
    partPath: "xl/worksheets/sheet1.xml",
    cells: new Map(),
    merges: [],
    comments: [],
    commentAuthors: [],
    columnWidths: new Map(),
    rowHeights: new Map(),
  };
}

function makeSnapshot(): XlsxSnapshot {
  const sheet = makeSheet("Sheet1");
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
  };
}

describe("xlsx:set-column-width", () => {
  it("stores the override in CSS px on Sheet.columnWidths", () => {
    const snap = makeSnapshot();
    const result = setColumnWidthHandler.apply(snap, { sheet: "Sheet1", column: 2, width: 140 });
    const sheet = result.next.root.sheets[0]!;
    expect(sheet.columnWidths.get(1)).toBe(140);
    expect(result.next.revision).toBe(1);
    expect(result.diff.changes).toHaveLength(1);
  });

  it("treats `null` width as a reset (delete the override)", () => {
    let snap = makeSnapshot();
    snap = setColumnWidthHandler.apply(snap, { sheet: "Sheet1", column: 2, width: 200 }).next;
    expect(snap.root.sheets[0]!.columnWidths.get(1)).toBe(200);
    snap = setColumnWidthHandler.apply(snap, { sheet: "Sheet1", column: 2, width: null }).next;
    expect(snap.root.sheets[0]!.columnWidths.has(1)).toBe(false);
  });

  it("rejects out-of-range widths and column indices", () => {
    const snap = makeSnapshot();
    expect(() => setColumnWidthHandler.apply(snap, { sheet: "Sheet1", column: 0, width: 100 })).toThrow(
      CommandError
    );
    expect(() => setColumnWidthHandler.apply(snap, { sheet: "Sheet1", column: 1, width: 4 })).toThrow(
      CommandError
    );
    expect(() => setColumnWidthHandler.apply(snap, { sheet: "Sheet1", column: 1, width: 5000 })).toThrow(
      CommandError
    );
  });

  it("is a no-op when the width matches the current override", () => {
    let snap = makeSnapshot();
    snap = setColumnWidthHandler.apply(snap, { sheet: "Sheet1", column: 1, width: 100 }).next;
    const r2 = setColumnWidthHandler.apply(snap, { sheet: "Sheet1", column: 1, width: 100 });
    expect(r2.diff.changes).toHaveLength(0);
  });
});

describe("xlsx:set-row-height", () => {
  it("stores the override in CSS px on Sheet.rowHeights", () => {
    const snap = makeSnapshot();
    const result = setRowHeightHandler.apply(snap, { sheet: "Sheet1", row: 3, height: 48 });
    const sheet = result.next.root.sheets[0]!;
    expect(sheet.rowHeights.get(2)).toBe(48);
    expect(result.diff.changes).toHaveLength(1);
  });

  it("treats `null` height as a reset", () => {
    let snap = makeSnapshot();
    snap = setRowHeightHandler.apply(snap, { sheet: "Sheet1", row: 2, height: 60 }).next;
    snap = setRowHeightHandler.apply(snap, { sheet: "Sheet1", row: 2, height: null }).next;
    expect(snap.root.sheets[0]!.rowHeights.has(1)).toBe(false);
  });

  it("rejects negative / zero / huge heights", () => {
    const snap = makeSnapshot();
    expect(() => setRowHeightHandler.apply(snap, { sheet: "Sheet1", row: 1, height: 0 })).toThrow(
      CommandError
    );
    expect(() => setRowHeightHandler.apply(snap, { sheet: "Sheet1", row: 1, height: -10 })).toThrow(
      CommandError
    );
    expect(() => setRowHeightHandler.apply(snap, { sheet: "Sheet1", row: 1, height: 9999 })).toThrow(
      CommandError
    );
  });
});
