import { describe, expect, it } from "vitest";
import {
  adjustForDeleteColumn,
  adjustForDeleteRow,
  adjustForInsertColumn,
  adjustForInsertRow,
  AbsRef,
  type CellRef,
} from "../references.js";
import { rewriteFormulaRefs } from "../rewrite-refs.js";

const ANCHOR: CellRef = { sheet: "Sheet1", row: 0, col: 1, abs: AbsRef.NONE };

describe("formula/rewrite-refs — insert", () => {
  it("shifts a single cell ref below the insertion point", () => {
    const out = rewriteFormulaRefs("A5", ANCHOR, (r) => adjustForInsertRow(r, "Sheet1", 2, 1));
    expect(out.text).toBe("A6");
    expect(out.changed).toBe(true);
    expect(out.hasRefError).toBe(false);
  });

  it("leaves a ref above the insertion point untouched", () => {
    const out = rewriteFormulaRefs("A1", ANCHOR, (r) => adjustForInsertRow(r, "Sheet1", 2, 1));
    expect(out.text).toBe("A1");
    expect(out.changed).toBe(false);
  });

  it("expands a range that straddles the insertion point", () => {
    const out = rewriteFormulaRefs("SUM(A1:A10)", ANCHOR, (r) => adjustForInsertRow(r, "Sheet1", 2, 1));
    expect(out.text).toBe("SUM(A1:A11)");
    expect(out.changed).toBe(true);
  });

  it("rewrites multiple refs in one expression", () => {
    // Insert one row at 0-based index 1 (= Excel row 2), so A1 stays
    // and every ref ≥ row 2 shifts down by one.
    const out = rewriteFormulaRefs("A1+SUM(A2:A10)", ANCHOR, (r) => adjustForInsertRow(r, "Sheet1", 1, 1));
    expect(out.text).toBe("A1+SUM(A3:A11)");
    expect(out.changed).toBe(true);
  });

  it("ignores refs in a different sheet", () => {
    const out = rewriteFormulaRefs("Sheet2!A5", ANCHOR, (r) => adjustForInsertRow(r, "Sheet1", 2, 1));
    expect(out.text).toBe("Sheet2!A5");
    expect(out.changed).toBe(false);
  });

  it("shifts a column ref past the insertion column", () => {
    const out = rewriteFormulaRefs("E1", ANCHOR, (r) => adjustForInsertColumn(r, "Sheet1", 2, 1));
    expect(out.text).toBe("F1");
    expect(out.changed).toBe(true);
  });
});

describe("formula/rewrite-refs — delete + #REF!", () => {
  it("rewrites a deleted cell ref to #REF!", () => {
    const out = rewriteFormulaRefs("A5", ANCHOR, (r) => adjustForDeleteRow(r, "Sheet1", 4, 1));
    expect(out.text).toBe("#REF!");
    expect(out.changed).toBe(true);
    expect(out.hasRefError).toBe(true);
  });

  it("rewrites a fully-consumed range to #REF!", () => {
    const out = rewriteFormulaRefs("SUM(A2:A5)", ANCHOR, (r) => adjustForDeleteRow(r, "Sheet1", 1, 5));
    expect(out.text).toBe("SUM(#REF!)");
    expect(out.hasRefError).toBe(true);
  });

  it("shifts a partially-overlapping range up", () => {
    const out = rewriteFormulaRefs("SUM(A1:A10)", ANCHOR, (r) => adjustForDeleteRow(r, "Sheet1", 4, 2));
    expect(out.text).toBe("SUM(A1:A8)");
    expect(out.changed).toBe(true);
    expect(out.hasRefError).toBe(false);
  });

  it("rewrites a column delete to #REF! when the cell is consumed", () => {
    const out = rewriteFormulaRefs("B1", ANCHOR, (r) => adjustForDeleteColumn(r, "Sheet1", 1, 1));
    expect(out.text).toBe("#REF!");
    expect(out.hasRefError).toBe(true);
  });

  it("preserves cross-sheet refs across a delete", () => {
    const out = rewriteFormulaRefs("Sheet2!A5", ANCHOR, (r) => adjustForDeleteRow(r, "Sheet1", 1, 5));
    expect(out.text).toBe("Sheet2!A5");
    expect(out.changed).toBe(false);
    expect(out.hasRefError).toBe(false);
  });
});
