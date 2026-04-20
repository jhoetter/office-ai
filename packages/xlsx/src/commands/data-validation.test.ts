import { describe, expect, it } from "vitest";
import { CommandError } from "@officeai/core";
import {
  addDataValidationHandler,
  removeDataValidationHandler,
  clearDataValidationsHandler,
} from "./data-validation.js";
import type { DataValidation, Sheet, XlsxSnapshot } from "../model/types.js";
import { defaultStyleTable } from "../model/style-table.js";

function makeSheet(rules: ReadonlyArray<DataValidation> = [], opaque?: string): Sheet {
  return {
    id: "sheet-1",
    sheetId: "1",
    name: "Sheet1",
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
    dataValidations: rules,
    tables: [],
    charts: [],
    ...(opaque ? { opaqueDataValidations: opaque } : {}),
  };
}

function makeSnapshot(rules: ReadonlyArray<DataValidation> = [], opaque?: string): XlsxSnapshot {
  const sheet = makeSheet(rules, opaque);
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

const RULE: DataValidation = {
  kind: "list",
  id: "dv-a",
  range: "A1:A10",
  source: "Yes,No,Maybe",
  formula: false,
  showDropDown: true,
  stopOnInvalid: true,
  allowBlank: true,
};

describe("xlsx:add-data-validation", () => {
  it("appends a typed list rule and marks the sheet dirty", () => {
    const snap = makeSnapshot();
    const out = addDataValidationHandler.apply(snap, { sheet: "Sheet1", rule: RULE });
    expect(out.next.root.sheets[0]!.dataValidations).toEqual([RULE]);
    expect(out.next.dirty.sheets.has("xl/worksheets/sheet1.xml")).toBe(true);
  });

  it("rejects duplicate ids", () => {
    const snap = makeSnapshot([RULE]);
    expect(() => addDataValidationHandler.apply(snap, { sheet: "Sheet1", rule: RULE })).toThrow(CommandError);
  });

  it("rejects empty source", () => {
    const snap = makeSnapshot();
    expect(() =>
      addDataValidationHandler.apply(snap, {
        sheet: "Sheet1",
        rule: { ...RULE, source: "" },
      })
    ).toThrow(CommandError);
  });
});

describe("xlsx:remove-data-validation", () => {
  it("drops the typed rule by id", () => {
    const snap = makeSnapshot([RULE]);
    const out = removeDataValidationHandler.apply(snap, { sheet: "Sheet1", id: RULE.id });
    expect(out.next.root.sheets[0]!.dataValidations).toEqual([]);
  });

  it("throws for unknown id", () => {
    const snap = makeSnapshot([RULE]);
    expect(() => removeDataValidationHandler.apply(snap, { sheet: "Sheet1", id: "nope" })).toThrow(
      CommandError
    );
  });
});

describe("xlsx:clear-data-validations", () => {
  it("wipes both typed and opaque rules", () => {
    const snap = makeSnapshot([RULE], '<dataValidations><dataValidation type="whole"/></dataValidations>');
    const out = clearDataValidationsHandler.apply(snap, { sheet: "Sheet1" });
    expect(out.next.root.sheets[0]!.dataValidations).toEqual([]);
    expect(out.next.root.sheets[0]!.opaqueDataValidations).toBeUndefined();
  });

  it("is a no-op when there's nothing to clear", () => {
    const snap = makeSnapshot();
    const out = clearDataValidationsHandler.apply(snap, { sheet: "Sheet1" });
    expect(out.diff.changes).toHaveLength(0);
  });
});
