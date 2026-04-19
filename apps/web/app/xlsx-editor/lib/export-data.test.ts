import { describe, expect, it } from "vitest";
import type { Cell, Sheet, XlsxSnapshot } from "@officeai/xlsx";
import {
  computeSheetBounds,
  sheetToCsv,
  sheetToTsv,
  workbookToJson,
} from "./export-data";

/* ── tiny snapshot factories ─────────────────────────────────────── */

interface CellSpec {
  readonly row: number;
  readonly col: number;
  readonly value: Cell["value"];
}

function makeSheet(name: string, cells: ReadonlyArray<CellSpec>, index = 0): Sheet {
  const map = new Map<string, Cell>();
  for (const c of cells) {
    map.set(`${c.row}:${c.col}`, {
      row: c.row,
      col: c.col,
      value: c.value,
    });
  }
  // The CSV/JSON helpers only touch `name`, `kind`, `index`, and
  // `cells.{values,get}` — cast covers the unused remainder of the
  // sheet schema.
  return {
    name,
    kind: "worksheet",
    index,
    cells: map,
  } as unknown as Sheet;
}

function makeWorkbook(sheets: ReadonlyArray<Sheet>): XlsxSnapshot {
  return { root: { sheets } } as unknown as XlsxSnapshot;
}

/* ── computeSheetBounds ──────────────────────────────────────────── */

describe("computeSheetBounds", () => {
  it("returns 0×0 for empty sheets", () => {
    expect(computeSheetBounds(makeSheet("S", []))).toEqual({ rows: 0, cols: 0 });
  });

  it("tracks the highest row and column with data", () => {
    const sheet = makeSheet("S", [
      { row: 0, col: 0, value: "a" },
      { row: 5, col: 2, value: 1 },
    ]);
    expect(computeSheetBounds(sheet)).toEqual({ rows: 6, cols: 3 });
  });
});

/* ── sheetToCsv / sheetToTsv ─────────────────────────────────────── */

describe("sheetToCsv", () => {
  it("emits a CRLF-terminated, comma-separated grid", () => {
    const sheet = makeSheet("S", [
      { row: 0, col: 0, value: "a" },
      { row: 0, col: 1, value: "b" },
      { row: 1, col: 0, value: 1 },
      { row: 1, col: 1, value: 2 },
    ]);
    expect(sheetToCsv(sheet)).toBe("a,b\r\n1,2\r\n");
  });

  it("returns an empty string for an empty sheet", () => {
    expect(sheetToCsv(makeSheet("S", []))).toBe("");
  });

  it("fills missing cells with empty fields up to the bounding box", () => {
    const sheet = makeSheet("S", [
      { row: 0, col: 0, value: "a" },
      { row: 1, col: 2, value: "z" },
    ]);
    expect(sheetToCsv(sheet)).toBe("a,,\r\n,,z\r\n");
  });

  it("quotes fields containing the delimiter", () => {
    const sheet = makeSheet("S", [
      { row: 0, col: 0, value: "a,b" },
      { row: 0, col: 1, value: "plain" },
    ]);
    expect(sheetToCsv(sheet)).toBe('"a,b",plain\r\n');
  });

  it("doubles embedded quotes and quotes the field", () => {
    const sheet = makeSheet("S", [{ row: 0, col: 0, value: 'say "hi"' }]);
    expect(sheetToCsv(sheet)).toBe('"say ""hi"""\r\n');
  });

  it("quotes fields containing newlines", () => {
    const sheet = makeSheet("S", [{ row: 0, col: 0, value: "line1\nline2" }]);
    expect(sheetToCsv(sheet)).toBe('"line1\nline2"\r\n');
  });

  it("renders booleans as TRUE/FALSE", () => {
    const sheet = makeSheet("S", [
      { row: 0, col: 0, value: true },
      { row: 0, col: 1, value: false },
    ]);
    expect(sheetToCsv(sheet)).toBe("TRUE,FALSE\r\n");
  });

  it("renders error sentinels by their Excel code", () => {
    const sheet = makeSheet("S", [
      { row: 0, col: 0, value: { kind: "error", code: "#REF!" } as Cell["value"] },
    ]);
    expect(sheetToCsv(sheet)).toBe("#REF!\r\n");
  });

  it("renders nulls as empty fields", () => {
    const sheet = makeSheet("S", [
      { row: 0, col: 0, value: null },
      { row: 0, col: 1, value: "x" },
    ]);
    expect(sheetToCsv(sheet)).toBe(",x\r\n");
  });
});

describe("sheetToTsv", () => {
  it("uses tabs and does not need to quote commas", () => {
    const sheet = makeSheet("S", [
      { row: 0, col: 0, value: "a,b" },
      { row: 0, col: 1, value: "c" },
    ]);
    expect(sheetToTsv(sheet)).toBe("a,b\tc\r\n");
  });

  it("still quotes embedded tabs", () => {
    const sheet = makeSheet("S", [{ row: 0, col: 0, value: "x\ty" }]);
    expect(sheetToTsv(sheet)).toBe('"x\ty"\r\n');
  });
});

/* ── workbookToJson ──────────────────────────────────────────────── */

describe("workbookToJson", () => {
  it("emits one entry per worksheet keyed by sheet name", () => {
    const wb = makeWorkbook([
      makeSheet(
        "First",
        [
          { row: 0, col: 0, value: "id" },
          { row: 0, col: 1, value: "qty" },
          { row: 1, col: 0, value: "abc" },
          { row: 1, col: 1, value: 42 },
        ],
        0
      ),
      makeSheet("Second", [{ row: 0, col: 0, value: true }], 1),
    ]);
    const parsed = JSON.parse(workbookToJson(wb)) as Record<
      string,
      Array<Record<string, unknown>>
    >;
    expect(Object.keys(parsed)).toEqual(["First", "Second"]);
    expect(parsed.First).toEqual([
      { A: "id", B: "qty" },
      { A: "abc", B: 42 },
    ]);
    expect(parsed.Second).toEqual([{ A: true }]);
  });

  it("uses Excel column letters past Z", () => {
    const sheet = makeSheet("S", [
      { row: 0, col: 25, value: "z" },
      { row: 0, col: 26, value: "aa" },
      { row: 0, col: 27, value: "ab" },
    ]);
    const wb = makeWorkbook([sheet]);
    const parsed = JSON.parse(workbookToJson(wb)) as Record<
      string,
      Array<Record<string, unknown>>
    >;
    expect(parsed.S[0]).toEqual({ Z: "z", AA: "aa", AB: "ab" });
  });

  it("encodes error sentinels as { error } objects", () => {
    const sheet = makeSheet("S", [
      { row: 0, col: 0, value: { kind: "error", code: "#DIV/0!" } as Cell["value"] },
    ]);
    const wb = makeWorkbook([sheet]);
    const parsed = JSON.parse(workbookToJson(wb)) as Record<
      string,
      Array<Record<string, unknown>>
    >;
    expect(parsed.S[0]).toEqual({ A: { error: "#DIV/0!" } });
  });

  it("skips empty rows", () => {
    const sheet = makeSheet("S", [
      { row: 0, col: 0, value: "x" },
      { row: 2, col: 0, value: "y" },
    ]);
    const wb = makeWorkbook([sheet]);
    const parsed = JSON.parse(workbookToJson(wb)) as Record<
      string,
      Array<Record<string, unknown>>
    >;
    expect(parsed.S).toEqual([{ A: "x" }, { A: "y" }]);
  });
});
