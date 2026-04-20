import { describe, expect, it } from "vitest";
import { CommandError } from "@officeai/core";
import { freezePanesHandler, unfreezePanesHandler } from "./freeze-panes.js";
import type { Sheet, XlsxSnapshot } from "../model/types.js";
import { defaultStyleTable } from "../model/style-table.js";

function makeSheet(name: string, freeze?: { rows: number; cols: number }): Sheet {
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
    hiddenRows: new Set(),
    hiddenCols: new Set(),
    images: [],
    conditionalFormats: [],
    opaqueConditionalFormats: [],
    dataValidations: [],
    tables: [],
    charts: [],
    ...(freeze ? { freeze } : {}),
  };
}

function makeSnapshot(freeze?: { rows: number; cols: number }): XlsxSnapshot {
  const sheet = makeSheet("Sheet1", freeze);
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

describe("xlsx:freeze-panes", () => {
  it("stores the freeze on the sheet and marks it dirty", () => {
    const snap = makeSnapshot();
    const result = freezePanesHandler.apply(snap, { sheet: "Sheet1", rows: 1, cols: 0 });
    const sheet = result.next.root.sheets[0]!;
    expect(sheet.freeze).toEqual({ rows: 1, cols: 0 });
    expect(result.next.dirty.sheets.has(sheet.partPath)).toBe(true);
    expect(result.diff.changes).toHaveLength(1);
  });

  it("treats {0,0} as an unfreeze", () => {
    let snap = makeSnapshot({ rows: 2, cols: 1 });
    snap = freezePanesHandler.apply(snap, { sheet: "Sheet1", rows: 0, cols: 0 }).next;
    expect(snap.root.sheets[0]!.freeze).toBeUndefined();
  });

  it("is a no-op when the freeze hasn't changed", () => {
    const snap = makeSnapshot({ rows: 1, cols: 0 });
    const result = freezePanesHandler.apply(snap, { sheet: "Sheet1", rows: 1, cols: 0 });
    expect(result.diff.changes).toHaveLength(0);
  });

  it("rejects negative / non-integer / oversized values", () => {
    const snap = makeSnapshot();
    expect(() => freezePanesHandler.apply(snap, { sheet: "Sheet1", rows: -1, cols: 0 })).toThrow(
      CommandError
    );
    expect(() => freezePanesHandler.apply(snap, { sheet: "Sheet1", rows: 0.5, cols: 0 })).toThrow(
      CommandError
    );
    expect(() => freezePanesHandler.apply(snap, { sheet: "Sheet1", rows: 0, cols: 16385 })).toThrow(
      CommandError
    );
  });
});

describe("xlsx:unfreeze-panes", () => {
  it("removes the freeze from the sheet", () => {
    const snap = makeSnapshot({ rows: 3, cols: 2 });
    const result = unfreezePanesHandler.apply(snap, { sheet: "Sheet1" });
    expect(result.next.root.sheets[0]!.freeze).toBeUndefined();
    expect(result.diff.changes).toHaveLength(1);
  });

  it("is a no-op when the sheet has no freeze", () => {
    const snap = makeSnapshot();
    const result = unfreezePanesHandler.apply(snap, { sheet: "Sheet1" });
    expect(result.diff.changes).toHaveLength(0);
  });
});
