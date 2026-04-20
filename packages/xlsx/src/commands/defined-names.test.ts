import { describe, expect, it } from "vitest";
import { CommandError } from "@officeai/core";
import {
  addDefinedNameHandler,
  removeDefinedNameHandler,
  updateDefinedNameHandler,
} from "./defined-names.js";
import type { DefinedName, Sheet, XlsxSnapshot } from "../model/types.js";
import { defaultStyleTable } from "../model/style-table.js";

function makeSheet(name: string): Sheet {
  return {
    id: `sheet-${name}`,
    sheetId: "1",
    name,
    index: 0,
    state: "visible",
    kind: "worksheet",
    partPath: `xl/worksheets/${name.toLowerCase()}.xml`,
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
  };
}

function makeSnapshot(initial: ReadonlyArray<DefinedName> = []): XlsxSnapshot {
  const sheet = makeSheet("Sheet1");
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
      definedNames: initial,
    },
    partHashes: {},
    container: undefined as never,
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

describe("xlsx:add-defined-name", () => {
  it("appends a workbook-scoped defined name and marks workbook dirty", () => {
    const snap = makeSnapshot();
    const result = addDefinedNameHandler.apply(snap, {
      name: "Revenue",
      refersTo: "Sheet1!$A$1:$A$10",
    });
    expect(result.next.root.definedNames).toHaveLength(1);
    expect(result.next.root.definedNames[0]).toMatchObject({
      name: "Revenue",
      refersTo: "Sheet1!$A$1:$A$10",
    });
    expect(result.next.dirty.workbook).toBe(true);
  });

  it("rejects invalid names", () => {
    const snap = makeSnapshot();
    expect(() => addDefinedNameHandler.apply(snap, { name: "1bad", refersTo: "A1" })).toThrow(CommandError);
    expect(() => addDefinedNameHandler.apply(snap, { name: "A1", refersTo: "B2" })).toThrow(CommandError);
    expect(() => addDefinedNameHandler.apply(snap, { name: "Has Space", refersTo: "B2" })).toThrow(
      CommandError
    );
  });

  it("rejects duplicates within the same scope", () => {
    let snap = makeSnapshot();
    snap = addDefinedNameHandler.apply(snap, { name: "X", refersTo: "A1" }).next;
    expect(() => addDefinedNameHandler.apply(snap, { name: "X", refersTo: "B1" })).toThrow(CommandError);
  });

  it("allows the same name in different scopes", () => {
    let snap = makeSnapshot();
    snap = addDefinedNameHandler.apply(snap, { name: "X", refersTo: "A1" }).next;
    const r2 = addDefinedNameHandler.apply(snap, {
      name: "X",
      refersTo: "B1",
      scope: "Sheet1",
    });
    expect(r2.next.root.definedNames).toHaveLength(2);
  });
});

describe("xlsx:update-defined-name", () => {
  it("renames and re-points an existing entry", () => {
    let snap = makeSnapshot();
    snap = addDefinedNameHandler.apply(snap, { name: "Old", refersTo: "A1" }).next;
    const r = updateDefinedNameHandler.apply(snap, {
      name: "Old",
      nextName: "New",
      refersTo: "B2",
    });
    expect(r.next.root.definedNames[0]).toMatchObject({ name: "New", refersTo: "B2" });
    expect(r.next.dirty.workbook).toBe(true);
  });

  it("rejects unknown names", () => {
    const snap = makeSnapshot();
    expect(() => updateDefinedNameHandler.apply(snap, { name: "Missing", refersTo: "A1" })).toThrow(
      CommandError
    );
  });
});

describe("xlsx:remove-defined-name", () => {
  it("drops a name and dirty's the workbook", () => {
    let snap = makeSnapshot();
    snap = addDefinedNameHandler.apply(snap, { name: "Doomed", refersTo: "A1" }).next;
    const r = removeDefinedNameHandler.apply(snap, { name: "Doomed" });
    expect(r.next.root.definedNames).toHaveLength(0);
    expect(r.next.dirty.workbook).toBe(true);
  });
});
