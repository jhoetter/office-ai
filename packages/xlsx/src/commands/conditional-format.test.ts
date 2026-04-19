import { describe, expect, it } from "vitest";
import { CommandError } from "@officeai/core";
import {
  addConditionalFormatHandler,
  removeConditionalFormatHandler,
  clearConditionalFormatsHandler,
} from "./conditional-format.js";
import type { ConditionalFormat, Sheet, XlsxSnapshot } from "../model/types.js";
import { defaultStyleTable } from "../model/style-table.js";

function makeSheet(rules: ReadonlyArray<ConditionalFormat> = []): Sheet {
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
    images: [],
    conditionalFormats: rules,
    opaqueConditionalFormats: [],
    dataValidations: [],
    tables: [],
    charts: [],
  };
}

function makeSnapshot(rules: ReadonlyArray<ConditionalFormat> = []): XlsxSnapshot {
  const sheet = makeSheet(rules);
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

const RULE_A: ConditionalFormat = {
  kind: "cellIs",
  id: "rule-a",
  range: "A1:A10",
  op: "gt",
  value: 5,
  overlay: { fill: "FFC7CE", fontColor: "9C0006" },
};

describe("xlsx:add-conditional-format", () => {
  it("appends a typed rule and marks the sheet dirty", () => {
    const snap = makeSnapshot();
    const out = addConditionalFormatHandler.apply(snap, {
      sheet: "Sheet1",
      rule: RULE_A,
    });
    expect(out.next.root.sheets[0]!.conditionalFormats).toEqual([RULE_A]);
    expect(out.next.dirty.sheets.has("xl/worksheets/sheet1.xml")).toBe(true);
    expect(out.diff.changes).toHaveLength(1);
  });

  it("rejects duplicate rule ids", () => {
    const snap = makeSnapshot([RULE_A]);
    expect(() => addConditionalFormatHandler.apply(snap, { sheet: "Sheet1", rule: RULE_A })).toThrow(
      CommandError
    );
  });

  it("rejects empty ids", () => {
    const snap = makeSnapshot();
    expect(() =>
      addConditionalFormatHandler.apply(snap, {
        sheet: "Sheet1",
        rule: { ...RULE_A, id: "" },
      })
    ).toThrow(CommandError);
  });
});

describe("xlsx:remove-conditional-format", () => {
  it("drops the named rule", () => {
    const snap = makeSnapshot([RULE_A]);
    const out = removeConditionalFormatHandler.apply(snap, {
      sheet: "Sheet1",
      id: RULE_A.id,
    });
    expect(out.next.root.sheets[0]!.conditionalFormats).toEqual([]);
  });

  it("throws for unknown ids", () => {
    const snap = makeSnapshot([RULE_A]);
    expect(() => removeConditionalFormatHandler.apply(snap, { sheet: "Sheet1", id: "nope" })).toThrow(
      CommandError
    );
  });
});

describe("xlsx:clear-conditional-formats", () => {
  it("empties the typed rule list", () => {
    const snap = makeSnapshot([RULE_A]);
    const out = clearConditionalFormatsHandler.apply(snap, { sheet: "Sheet1" });
    expect(out.next.root.sheets[0]!.conditionalFormats).toEqual([]);
  });

  it("is a no-op when there are no typed rules", () => {
    const snap = makeSnapshot();
    const out = clearConditionalFormatsHandler.apply(snap, { sheet: "Sheet1" });
    expect(out.diff.changes).toHaveLength(0);
  });

  it("leaves opaque (imported) rules alone", () => {
    const snap = makeSnapshot([RULE_A]);
    // Inject opaque blocks the same way the parser would.
    const sheet = snap.root.sheets[0]!;
    const sheets = snap.root.sheets.slice();
    sheets[0] = {
      ...sheet,
      opaqueConditionalFormats: [
        '<conditionalFormatting sqref="B1:B10"><cfRule type="cellIs" priority="1"/></conditionalFormatting>',
      ],
    };
    const snap2: XlsxSnapshot = {
      ...snap,
      root: { ...snap.root, sheets },
    };
    const out = clearConditionalFormatsHandler.apply(snap2, { sheet: "Sheet1" });
    expect(out.next.root.sheets[0]!.conditionalFormats).toEqual([]);
    expect(out.next.root.sheets[0]!.opaqueConditionalFormats).toHaveLength(1);
  });
});
