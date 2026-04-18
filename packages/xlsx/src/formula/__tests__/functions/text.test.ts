import { describe, expect, it } from "vitest";
import { evaluate } from "../../evaluator.js";
import { createRegistry, type EvalContext } from "../../function-registry.js";
import { parse } from "../../parser.js";
import { AbsRef, type CellRef } from "../../references.js";
import { registerText } from "../../functions/text.js";
import { bool, isError, num, str, type Value } from "../../values.js";

const ANCHOR: CellRef = { sheet: "S", row: 0, col: 0, abs: AbsRef.NONE };

function makeCtx(): EvalContext {
  const reg = createRegistry();
  registerText(reg);
  return {
    getCell: () => num(0),
    getRange: () => [[num(0)]],
    resolveName: () => undefined,
    now: () => 0,
    random: () => 0,
    anchor: ANCHOR,
    registry: reg,
  };
}

const ctx = makeCtx();
const ev = (text: string): Value => evaluate(parse(text, { anchor: ANCHOR }).ast, ctx);

describe("formula/functions/text — CONCATENATE", () => {
  it("joins string scalars", () => {
    expect(ev('=CONCATENATE("a","b","c")')).toEqual(str("abc"));
  });
  it("coerces numbers and booleans", () => {
    expect(ev('=CONCATENATE("v=",1,"/",TRUE)')).toEqual(str("v=1/TRUE"));
  });
  it("rejects multi-cell ranges with #VALUE!", () => {
    expect(ev("=CONCATENATE({1,2})")).toMatchObject({ kind: "e" });
  });
  it("propagates errors from inputs", () => {
    expect(ev('=CONCATENATE("x", #N/A)')).toMatchObject({ kind: "e" });
  });
});

describe("formula/functions/text — CONCAT", () => {
  it("concatenates scalar args", () => {
    expect(ev('=CONCAT("a","b","c")')).toEqual(str("abc"));
  });
  it("flattens 2-D ranges row-major", () => {
    expect(ev("=CONCAT({1,2;3,4})")).toEqual(str("1234"));
  });
  it("propagates errors inside a range", () => {
    expect(ev("=CONCAT({1,#REF!})")).toMatchObject({ kind: "e" });
  });
});

describe("formula/functions/text — TEXTJOIN", () => {
  it("joins with delimiter", () => {
    expect(ev('=TEXTJOIN(",",TRUE,"a","b","c")')).toEqual(str("a,b,c"));
  });
  it("skips empty strings when ignoreEmpty=TRUE", () => {
    expect(ev('=TEXTJOIN("-",TRUE,"a","","b")')).toEqual(str("a-b"));
  });
  it("includes empty strings when ignoreEmpty=FALSE", () => {
    expect(ev('=TEXTJOIN("-",FALSE,"a","","b")')).toEqual(str("a--b"));
  });
  it("flattens range arguments", () => {
    expect(ev('=TEXTJOIN(",",TRUE,{1,2;3,4})')).toEqual(str("1,2,3,4"));
  });
});

describe("formula/functions/text — LEFT", () => {
  it("defaults to one char", () => {
    expect(ev('=LEFT("Hello")')).toEqual(str("H"));
  });
  it("returns the requested prefix", () => {
    expect(ev('=LEFT("Hello", 3)')).toEqual(str("Hel"));
  });
  it("clamps n > length", () => {
    expect(ev('=LEFT("ab", 10)')).toEqual(str("ab"));
  });
  it("rejects negative n", () => {
    expect(ev('=LEFT("abc", -1)')).toMatchObject({ kind: "e" });
  });
});

describe("formula/functions/text — RIGHT", () => {
  it("defaults to one char", () => {
    expect(ev('=RIGHT("Hello")')).toEqual(str("o"));
  });
  it("returns the requested suffix", () => {
    expect(ev('=RIGHT("Hello", 3)')).toEqual(str("llo"));
  });
  it("clamps n > length", () => {
    expect(ev('=RIGHT("abc", 100)')).toEqual(str("abc"));
  });
  it("rejects negative n", () => {
    expect(ev('=RIGHT("abc", -1)')).toMatchObject({ kind: "e" });
  });
});

describe("formula/functions/text — MID", () => {
  it("returns a substring (1-based)", () => {
    expect(ev('=MID("abcdef", 2, 3)')).toEqual(str("bcd"));
  });
  it("returns empty when start past end", () => {
    expect(ev('=MID("abc", 10, 3)')).toEqual(str(""));
  });
  it("rejects start < 1", () => {
    expect(ev('=MID("abc", 0, 3)')).toMatchObject({ kind: "e" });
  });
  it("rejects negative length", () => {
    expect(ev('=MID("abc", 1, -1)')).toMatchObject({ kind: "e" });
  });
});

describe("formula/functions/text — LEN", () => {
  it("counts characters", () => {
    expect(ev('=LEN("hello")')).toEqual(num(5));
  });
  it("returns 0 for empty string", () => {
    expect(ev('=LEN("")')).toEqual(num(0));
  });
  it("coerces numbers to text first", () => {
    expect(ev("=LEN(1234)")).toEqual(num(4));
  });
});

describe("formula/functions/text — TRIM", () => {
  it("trims leading and trailing whitespace", () => {
    expect(ev('=TRIM("  hi  ")')).toEqual(str("hi"));
  });
  it("collapses internal runs", () => {
    expect(ev('=TRIM("a   b   c")')).toEqual(str("a b c"));
  });
  it("returns empty for whitespace-only", () => {
    expect(ev('=TRIM("    ")')).toEqual(str(""));
  });
});

describe("formula/functions/text — UPPER / LOWER / PROPER", () => {
  it("UPPER", () => {
    expect(ev('=UPPER("hi there")')).toEqual(str("HI THERE"));
  });
  it("LOWER", () => {
    expect(ev('=LOWER("HI THERE")')).toEqual(str("hi there"));
  });
  it("PROPER capitalises each word", () => {
    expect(ev('=PROPER("hello WORLD foo")')).toEqual(str("Hello World Foo"));
  });
  it("PROPER handles non-letter separators", () => {
    expect(ev('=PROPER("a-b c.d")')).toEqual(str("A-B C.D"));
  });
});

describe("formula/functions/text — FIND", () => {
  it("finds a substring", () => {
    expect(ev('=FIND("a", "banana")')).toEqual(num(2));
  });
  it("is case-sensitive", () => {
    expect(ev('=FIND("A", "banana")')).toMatchObject({ kind: "e" });
  });
  it("respects start position", () => {
    expect(ev('=FIND("a", "banana", 3)')).toEqual(num(4));
  });
  it("rejects start < 1", () => {
    expect(ev('=FIND("a", "banana", 0)')).toMatchObject({ kind: "e" });
  });
});

describe("formula/functions/text — SEARCH", () => {
  it("is case-insensitive", () => {
    expect(ev('=SEARCH("A", "banana")')).toEqual(num(2));
  });
  it("supports the * wildcard", () => {
    expect(ev('=SEARCH("a*c", "abc")')).toEqual(num(1));
  });
  it("supports the ? wildcard", () => {
    expect(ev('=SEARCH("b?n", "banana")')).toEqual(num(1));
  });
  it("supports ~ to escape wildcards", () => {
    expect(ev('=SEARCH("a~?c", "xa?cx")')).toEqual(num(2));
  });
  it("returns #VALUE! on miss", () => {
    expect(ev('=SEARCH("zzz", "abc")')).toMatchObject({ kind: "e" });
  });
});

describe("formula/functions/text — SUBSTITUTE", () => {
  it("replaces every occurrence by default", () => {
    expect(ev('=SUBSTITUTE("banana", "a", "o")')).toEqual(str("bonono"));
  });
  it("replaces only the nth occurrence", () => {
    expect(ev('=SUBSTITUTE("banana", "a", "o", 2)')).toEqual(str("banona"));
  });
  it("returns text unchanged if old is empty", () => {
    expect(ev('=SUBSTITUTE("hi", "", "x")')).toEqual(str("hi"));
  });
  it("returns text unchanged if nth occurrence missing", () => {
    expect(ev('=SUBSTITUTE("abc", "a", "x", 5)')).toEqual(str("abc"));
  });
});

describe("formula/functions/text — REPLACE", () => {
  it("replaces by position+length", () => {
    expect(ev('=REPLACE("abcdef", 2, 3, "ZZ")')).toEqual(str("aZZef"));
  });
  it("inserts when length=0", () => {
    expect(ev('=REPLACE("abc", 2, 0, "X")')).toEqual(str("aXbc"));
  });
  it("rejects start < 1", () => {
    expect(ev('=REPLACE("abc", 0, 1, "x")')).toMatchObject({ kind: "e" });
  });
  it("rejects negative length", () => {
    expect(ev('=REPLACE("abc", 1, -1, "x")')).toMatchObject({ kind: "e" });
  });
});

describe("formula/functions/text — REPT", () => {
  it("repeats a string n times", () => {
    expect(ev('=REPT("ab", 3)')).toEqual(str("ababab"));
  });
  it("returns empty when n=0", () => {
    expect(ev('=REPT("ab", 0)')).toEqual(str(""));
  });
  it("rejects negative n", () => {
    expect(ev('=REPT("a", -1)')).toMatchObject({ kind: "e" });
  });
  it("rejects results > 32767 chars", () => {
    expect(ev('=REPT("a", 32768)')).toMatchObject({ kind: "e" });
  });
});

describe("formula/functions/text — TEXT (P0 minimal)", () => {
  it('formats integers with "0"', () => {
    expect(ev('=TEXT(42, "0")')).toEqual(str("42"));
  });
  it('formats with fixed decimals "0.00"', () => {
    expect(ev('=TEXT(1.5, "0.00")')).toEqual(str("1.50"));
  });
  it('formats with thousands separator "#,##0"', () => {
    expect(ev('=TEXT(1234567, "#,##0")')).toEqual(str("1,234,567"));
  });
  it('formats currency "$#,##0.00"', () => {
    expect(ev('=TEXT(1234.5, "$#,##0.00")')).toEqual(str("$1,234.50"));
  });
  it('formats percent "0.0%"', () => {
    expect(ev('=TEXT(0.125, "0.0%")')).toEqual(str("12.5%"));
  });
  it("falls back gracefully for unknown formats", () => {
    expect(ev('=TEXT(42, "weird")')).toEqual(str("42"));
  });
});

describe("formula/functions/text — VALUE", () => {
  it("parses a numeric string", () => {
    expect(ev('=VALUE("42")')).toEqual(num(42));
  });
  it("parses a decimal string", () => {
    expect(ev('=VALUE("3.14")')).toEqual(num(3.14));
  });
  it("returns #VALUE! on garbage", () => {
    expect(ev('=VALUE("nope")')).toMatchObject({ kind: "e" });
  });
});

describe("formula/functions/text — NUMBERVALUE", () => {
  it("uses default separators", () => {
    expect(ev('=NUMBERVALUE("1,234.5")')).toEqual(num(1234.5));
  });
  it("respects custom decimal separator", () => {
    expect(ev('=NUMBERVALUE("1.234,5", ",", ".")')).toEqual(num(1234.5));
  });
  it("returns #VALUE! on garbage", () => {
    expect(ev('=NUMBERVALUE("abc")')).toMatchObject({ kind: "e" });
  });
});

describe("formula/functions/text — CHAR / CODE", () => {
  it("CHAR: code → char", () => {
    expect(ev("=CHAR(65)")).toEqual(str("A"));
  });
  it("CHAR: out-of-range → #VALUE!", () => {
    expect(ev("=CHAR(0)")).toMatchObject({ kind: "e" });
    expect(ev("=CHAR(256)")).toMatchObject({ kind: "e" });
  });
  it("CODE: first char → code", () => {
    expect(ev('=CODE("ABC")')).toEqual(num(65));
  });
  it("CODE: empty → #VALUE!", () => {
    expect(ev('=CODE("")')).toMatchObject({ kind: "e" });
  });
});

describe("formula/functions/text — EXACT", () => {
  it("matches identical strings", () => {
    expect(ev('=EXACT("abc", "abc")')).toEqual(bool(true));
  });
  it("is case-sensitive", () => {
    expect(ev('=EXACT("abc", "ABC")')).toEqual(bool(false));
  });
  it("returns false for differing strings", () => {
    expect(ev('=EXACT("a", "b")')).toEqual(bool(false));
  });
});

describe("formula/functions/text — T", () => {
  it("returns text unchanged", () => {
    expect(ev('=T("hello")')).toEqual(str("hello"));
  });
  it("returns empty for numbers", () => {
    expect(ev("=T(42)")).toEqual(str(""));
  });
  it("returns empty for booleans", () => {
    expect(ev("=T(TRUE)")).toEqual(str(""));
  });
  it("propagates errors", () => {
    const out = ev("=T(#N/A)");
    expect(isError(out)).toBe(true);
  });
});
