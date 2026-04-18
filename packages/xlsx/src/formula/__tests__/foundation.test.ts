import { describe, expect, it } from "vitest";
import { ErrorKinds, Errors, parseErrorLiteral, refWithCycle, refWithDeletedTarget } from "../errors.js";
import {
  add,
  bool,
  compare,
  concat,
  div,
  eq,
  gt,
  lt,
  mul,
  neg,
  num,
  pct,
  pow,
  rangeVal,
  str,
  sub,
  toBoolean,
  toNumber,
  toString,
  type Value,
} from "../values.js";
import {
  AbsRef,
  adjustForDeleteColumn,
  adjustForDeleteRow,
  adjustForInsertColumn,
  adjustForInsertRow,
  a1ToR1C1,
  cellRefKey,
  colLetterToIndex,
  indexToColLetter,
  isRefError,
  parseA1,
  parseA1Range,
  parseCellKey,
  r1c1ToA1,
  serializeCellRef,
  serializeRangeRef,
  stripSheetPrefix,
} from "../references.js";

describe("formula/errors", () => {
  it("interns one singleton per kind", () => {
    expect(Errors[ErrorKinds.DIV0]).toBe(Errors[ErrorKinds.DIV0]);
    expect(Errors[ErrorKinds.NAME].kind).toBe("#NAME?");
  });

  it("recognises all canonical error literals", () => {
    expect(parseErrorLiteral("#DIV/0!")?.kind).toBe(ErrorKinds.DIV0);
    expect(parseErrorLiteral("#NAME?")?.kind).toBe(ErrorKinds.NAME);
    expect(parseErrorLiteral("#REF!")?.kind).toBe(ErrorKinds.REF);
    expect(parseErrorLiteral("#GETTING_DATA")?.kind).toBe(ErrorKinds.GETTING_DATA);
    expect(parseErrorLiteral("not-an-error")).toBeUndefined();
  });

  it("carries cycle metadata on refWithCycle", () => {
    const e = refWithCycle(["S!0:0", "S!1:1"]);
    expect(e.kind).toBe(ErrorKinds.REF);
    expect((e.meta as { cycle: string[] }).cycle).toEqual(["S!0:0", "S!1:1"]);
  });

  it("carries deletedRef metadata on refWithDeletedTarget", () => {
    const e = refWithDeletedTarget("S!A1");
    expect((e.meta as { deletedRef: string }).deletedRef).toBe("S!A1");
  });
});

describe("formula/values — coercion", () => {
  it("toNumber coerces booleans, strings, and blanks", () => {
    expect(toNumber(bool(true))).toEqual(num(1));
    expect(toNumber(bool(false))).toEqual(num(0));
    expect(toNumber(str("3.14"))).toEqual(num(3.14));
    expect(toNumber(str(""))).toEqual(num(0));
    expect((toNumber(str("nope")) as Value).kind).toBe("e");
  });

  it("toString of number uses Excel general format", () => {
    expect(toString(num(42))).toEqual(str("42"));
    expect(toString(num(0.1 + 0.2))).toEqual(str("0.3"));
    expect(toString(bool(true))).toEqual(str("TRUE"));
  });

  it("toBoolean accepts only canonical strings", () => {
    expect(toBoolean(str("TRUE"))).toEqual(bool(true));
    expect(toBoolean(str("false"))).toEqual(bool(false));
    expect((toBoolean(str("nope")) as Value).kind).toBe("e");
  });

  it("propagates errors through coercion", () => {
    const e: Value = { kind: "e", v: Errors[ErrorKinds.DIV0] };
    expect(toNumber(e)).toBe(e);
    expect(toString(e)).toBe(e);
    expect(toBoolean(e)).toBe(e);
  });
});

describe("formula/values — arithmetic", () => {
  it("add/sub/mul/div handle numbers and propagate errors", () => {
    expect(add(num(2), num(3))).toEqual(num(5));
    expect(sub(num(10), num(4))).toEqual(num(6));
    expect(mul(num(3), num(4))).toEqual(num(12));
    expect(div(num(10), num(4))).toEqual(num(2.5));
    expect(div(num(1), num(0))).toEqual({ kind: "e", v: Errors[ErrorKinds.DIV0] });
  });

  it("pow with non-finite returns #NUM!", () => {
    expect(pow(num(0), num(-1))).toEqual({ kind: "e", v: Errors[ErrorKinds.NUM] });
    expect(pow(num(2), num(3))).toEqual(num(8));
    expect(pow(num(2), num(10))).toEqual(num(1024));
  });

  it("neg / pct", () => {
    expect(neg(num(7))).toEqual(num(-7));
    expect(pct(num(50))).toEqual(num(0.5));
  });

  it("concat coerces both sides", () => {
    expect(concat(str("a"), num(1))).toEqual(str("a1"));
    expect(concat(bool(true), str("X"))).toEqual(str("TRUEX"));
  });
});

describe("formula/values — comparison", () => {
  it("number<string<bool cross-type ordering", () => {
    expect(compare(num(1), str("a"))).toBe(-1);
    expect(compare(str("a"), bool(false))).toBe(-1);
    expect(compare(num(1), bool(false))).toBe(-1);
  });

  it("strings are case-insensitive", () => {
    expect(compare(str("A"), str("a"))).toBe(0);
    expect(compare(str("A"), str("B"))).toBe(-1);
  });

  it("eq/neq/lt/gt return BoolValue", () => {
    expect(eq(num(1), num(1))).toEqual(bool(true));
    expect(lt(num(1), num(2))).toEqual(bool(true));
    expect(gt(num(2), num(1))).toEqual(bool(true));
  });

  it("collapses 1x1 ranges to scalars", () => {
    const r = rangeVal([[num(7)]]);
    expect(compare(r, num(7))).toBe(0);
  });
});

describe("formula/references — A1 parse / serialize", () => {
  it("colLetterToIndex round-trips up to ZZZ", () => {
    expect(colLetterToIndex("A")).toBe(0);
    expect(colLetterToIndex("Z")).toBe(25);
    expect(colLetterToIndex("AA")).toBe(26);
    expect(colLetterToIndex("XFD")).toBe(16383);
    expect(indexToColLetter(0)).toBe("A");
    expect(indexToColLetter(26)).toBe("AA");
    expect(indexToColLetter(16383)).toBe("XFD");
  });

  it("parses simple A1 with default sheet", () => {
    const r = parseA1("B3", "Sheet1");
    expect(r).toEqual({ sheet: "Sheet1", row: 2, col: 1, abs: AbsRef.NONE });
  });

  it("parses absolute markers", () => {
    expect(parseA1("$A1", "S")?.abs).toBe(AbsRef.COLUMN);
    expect(parseA1("A$1", "S")?.abs).toBe(AbsRef.ROW);
    expect(parseA1("$A$1", "S")?.abs).toBe(AbsRef.ALL);
  });

  it("parses sheet-qualified refs (bare and quoted)", () => {
    expect(parseA1("Other!A1", "S")?.sheet).toBe("Other");
    expect(parseA1("'My Sheet'!A1", "S")?.sheet).toBe("My Sheet");
    expect(parseA1("'It''s'!A1", "S")?.sheet).toBe("It's");
  });

  it("rejects garbage", () => {
    expect(parseA1("", "S")).toBeUndefined();
    expect(parseA1("A0", "S")).toBeUndefined();
    expect(parseA1("XFE1", "S")).toBeUndefined();
  });

  it("parses ranges including whole-column and whole-row", () => {
    const r = parseA1Range("A1:B5", "S")!;
    expect(r).toMatchObject({ sheet: "S", r0: 0, c0: 0, r1: 4, c1: 1 });
    expect(parseA1Range("A:A", "S")?.c1).toBe(0);
    expect(parseA1Range("3:5", "S")?.r1).toBe(4);
  });

  it("normalises out-of-order ranges", () => {
    const r = parseA1Range("B5:A1", "S")!;
    expect(r.r0).toBe(0);
    expect(r.c0).toBe(0);
    expect(r.r1).toBe(4);
    expect(r.c1).toBe(1);
  });

  it("serializes refs with anchor-aware sheet omission", () => {
    const ref = parseA1("A1", "S")!;
    expect(serializeCellRef(ref, { sheet: "S" })).toBe("A1");
    expect(serializeCellRef(ref)).toBe("S!A1");
    expect(serializeCellRef(parseA1("$A$1", "S")!)).toBe("S!$A$1");
  });

  it("serializes ranges", () => {
    const r = parseA1Range("$A$1:$B$5", "S")!;
    expect(serializeRangeRef(r, { sheet: "S" })).toBe("$A$1:$B$5");
  });

  it("quotes sheet names that contain special chars", () => {
    const ref = parseA1("'My Sheet'!A1", "S")!;
    expect(serializeCellRef(ref)).toBe("'My Sheet'!A1");
  });
});

describe("formula/references — A1 ↔ R1C1", () => {
  it("converts relative refs around an anchor", () => {
    const anchor = parseA1("B2", "S")!;
    const ref = parseA1("C3", "S")!;
    expect(a1ToR1C1(ref, anchor)).toBe("R[1]C[1]");
    const back = r1c1ToA1("R[1]C[1]", anchor)!;
    expect(back.row).toBe(2);
    expect(back.col).toBe(2);
  });

  it("converts absolute refs", () => {
    const anchor = parseA1("B2", "S")!;
    const ref = parseA1("$A$1", "S")!;
    expect(a1ToR1C1(ref, anchor)).toBe("R1C1");
    expect(r1c1ToA1("R1C1", anchor)?.row).toBe(0);
  });
});

describe("formula/references — insert/delete adjust", () => {
  const ref = parseA1("B5", "S")!;

  it("insert row above shifts down", () => {
    const out = adjustForInsertRow(ref, "S", 2, 3);
    expect(out).toMatchObject({ row: 7 });
  });

  it("insert row below leaves alone", () => {
    const out = adjustForInsertRow(ref, "S", 6, 3);
    expect(out).toMatchObject({ row: 4 });
  });

  it("delete-row consuming cell yields #REF!", () => {
    const out = adjustForDeleteRow(ref, "S", 4, 1);
    expect(isRefError(out)).toBe(true);
  });

  it("delete-row above shifts up", () => {
    const out = adjustForDeleteRow(ref, "S", 1, 2);
    expect(out).toMatchObject({ row: 2 });
  });

  it("range partially overlapped by delete shrinks", () => {
    const r = parseA1Range("A1:A10", "S")!;
    const out = adjustForDeleteRow(r, "S", 3, 2);
    expect(out).toMatchObject({ r0: 0, r1: 7 });
  });

  it("range fully consumed returns #REF!", () => {
    const r = parseA1Range("A2:A4", "S")!;
    const out = adjustForDeleteRow(r, "S", 1, 5);
    expect(isRefError(out)).toBe(true);
  });

  it("column adjust mirrors row adjust", () => {
    const r = parseA1("E5", "S")!;
    expect(adjustForInsertColumn(r, "S", 2, 3)).toMatchObject({ col: 7 });
    expect(adjustForDeleteColumn(r, "S", 1, 2)).toMatchObject({ col: 2 });
    expect(isRefError(adjustForDeleteColumn(r, "S", 4, 1))).toBe(true);
  });

  it("ignores refs on a different sheet", () => {
    const out = adjustForInsertRow(ref, "Other", 0, 5);
    expect(out).toBe(ref);
  });
});

describe("formula/references — sheet prefix + cell key", () => {
  it("strips bare and quoted sheet prefixes", () => {
    expect(stripSheetPrefix("Sheet1!A1")).toEqual({ sheet: "Sheet1", rest: "A1" });
    expect(stripSheetPrefix("'My Sheet'!A1")).toEqual({ sheet: "My Sheet", rest: "A1" });
    expect(stripSheetPrefix("A1")).toBeUndefined();
  });

  it("cellRefKey + parseCellKey round-trip", () => {
    const ref = parseA1("'My Sheet'!C7", "S")!;
    const k = cellRefKey(ref);
    expect(k).toBe("My Sheet!6:2");
    expect(parseCellKey(k)).toEqual({ sheet: "My Sheet", row: 6, col: 2 });
  });
});
