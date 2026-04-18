import { describe, expect, it } from "vitest";
import {
  delimitedToSnapshot,
  sniffDelimiter,
  snapshotToTsv,
  tsvToSnapshot,
  type XlsxClipboardSnapshot,
} from "./snapshot.js";

const literalSnap = (
  rows: ReadonlyArray<ReadonlyArray<string | number | boolean | null>>
): XlsxClipboardSnapshot => ({
  origin: { sheet: "x", range: "A1" },
  width: rows[0]?.length ?? 0,
  height: rows.length,
  cells: rows.map((row) => row.map((v) => (v === null ? null : { value: v }))),
  merges: [],
});

describe("snapshotToTsv", () => {
  it("emits tab-separated rows", () => {
    expect(
      snapshotToTsv(
        literalSnap([
          ["a", 1],
          ["b", 2],
        ])
      )
    ).toBe("a\t1\nb\t2");
  });

  it("renders booleans as TRUE/FALSE", () => {
    expect(snapshotToTsv(literalSnap([[true, false]]))).toBe("TRUE\tFALSE");
  });

  it("renders empty cells as empty strings", () => {
    expect(snapshotToTsv(literalSnap([["a", null, "c"]]))).toBe("a\t\tc");
  });

  it("emits formulas with leading =", () => {
    const snap: XlsxClipboardSnapshot = {
      origin: { sheet: "x", range: "A1" },
      width: 1,
      height: 1,
      cells: [[{ value: 5, formula: "A1+B1" }]],
      merges: [],
    };
    expect(snapshotToTsv(snap)).toBe("=A1+B1");
  });

  it("collapses embedded tabs / newlines to spaces", () => {
    expect(snapshotToTsv(literalSnap([["a\tb\nc"]]))).toBe("a b c");
  });
});

describe("tsvToSnapshot", () => {
  it("round-trips a simple grid", () => {
    const snap = tsvToSnapshot("1\t2\n3\t4");
    expect(snap.height).toBe(2);
    expect(snap.width).toBe(2);
    expect(snap.cells[0]?.[0]?.value).toBe(1);
    expect(snap.cells[1]?.[1]?.value).toBe(4);
  });

  it("decodes booleans + leading-= as formulas", () => {
    const snap = tsvToSnapshot("=A1+B1\ttrue\tFALSE");
    expect(snap.cells[0]?.[0]?.formula).toBe("A1+B1");
    expect(snap.cells[0]?.[1]?.value).toBe(true);
    expect(snap.cells[0]?.[2]?.value).toBe(false);
  });

  it("drops a single trailing empty row (Excel adds one)", () => {
    expect(tsvToSnapshot("a\nb\n").height).toBe(2);
  });

  it("preserves empty cells as null entries", () => {
    const snap = tsvToSnapshot("a\t\tc");
    expect(snap.width).toBe(3);
    expect(snap.cells[0]?.[1]).toBeNull();
  });
});

describe("sniffDelimiter", () => {
  it("prefers tab when present", () => {
    expect(sniffDelimiter("a\tb,c")).toBe("\t");
  });

  it("picks comma over semicolon when comma is more frequent", () => {
    expect(sniffDelimiter("a,b,c;d")).toBe(",");
  });

  it("picks semicolon when it dominates", () => {
    expect(sniffDelimiter("a;b;c;d,e")).toBe(";");
  });

  it("falls back to tab on a single-cell payload", () => {
    expect(sniffDelimiter("hello")).toBe("\t");
  });
});

describe("delimitedToSnapshot", () => {
  it("parses CSV with quoted fields and embedded commas", () => {
    const snap = delimitedToSnapshot('"a,b",c\n"d","e,""f"', ",");
    expect(snap.height).toBe(2);
    expect(snap.cells[0]?.[0]?.value).toBe("a,b");
    expect(snap.cells[0]?.[1]?.value).toBe("c");
    expect(snap.cells[1]?.[1]?.value).toBe('e,"f');
  });

  it("respects newlines inside quoted fields", () => {
    const snap = delimitedToSnapshot('"line1\nline2",b\nc,d', ",");
    expect(snap.height).toBe(2);
    expect(snap.cells[0]?.[0]?.value).toBe("line1\nline2");
    expect(snap.cells[1]?.[0]?.value).toBe("c");
  });

  it("supports semicolon delimiter (German Excel CSV)", () => {
    const snap = delimitedToSnapshot("1;2;3\n4;5;6", ";");
    expect(snap.height).toBe(2);
    expect(snap.width).toBe(3);
    expect(snap.cells[1]?.[2]?.value).toBe(6);
  });
});
