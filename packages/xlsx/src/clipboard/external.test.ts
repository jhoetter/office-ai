import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { describe, expect, it } from "vitest";
import { parseExternalClipboard, parseFingerprintHtml, parseHtmlTable } from "./external.js";
import type { XlsxClipboardSnapshot } from "./snapshot.js";

const here = dirname(fileURLToPath(import.meta.url));
const fixture = (name: string): string => readFileSync(join(here, "__fixtures__", name), "utf8");

describe("parseHtmlTable — Excel desktop fixture", () => {
  const snap = parseHtmlTable(fixture("excel-desktop.html"))!;

  it("captures the data rows and ignores Excel chrome", () => {
    expect(snap).not.toBeNull();
    expect(snap.height).toBe(3);
    expect(snap.width).toBe(4);
    expect(snap.cells[0]?.[0]?.value).toBe("Region");
    expect(snap.cells[1]?.[0]?.value).toBe("EMEA");
    expect(snap.cells[2]?.[0]?.value).toBe("APAC");
  });

  it("converts numeric cells to numbers", () => {
    expect(snap.cells[1]?.[2]?.value).toBe(1240);
    expect(snap.cells[1]?.[3]?.value).toBe(54980.5);
    expect(snap.cells[2]?.[2]?.value).toBe(980);
    expect(snap.cells[2]?.[3]?.value).toBe(43210.75);
  });
});

describe("parseHtmlTable — Google Sheets fixture", () => {
  const snap = parseHtmlTable(fixture("google-sheets.html"))!;

  it("ignores the wrapper sheets-html-origin element and grabs cell text", () => {
    expect(snap.height).toBe(3);
    expect(snap.width).toBe(3);
    expect(snap.cells[0]?.map((c) => c?.value)).toEqual(["Item", "Qty", "Price"]);
    expect(snap.cells[1]?.map((c) => c?.value)).toEqual(["Widget", 12, 9.99]);
    expect(snap.cells[2]?.map((c) => c?.value)).toEqual(["Gadget", 5, 24.5]);
  });
});

describe("parseHtmlTable — Apple Numbers fixture", () => {
  const snap = parseHtmlTable(fixture("numbers.html"))!;

  it("treats date-shaped strings as strings (not numbers)", () => {
    expect(snap.cells[1]?.[0]?.value).toBe("2026-04-15");
    expect(snap.cells[3]?.[0]?.value).toBe("2026-04-17");
  });

  it("captures the score column as numbers", () => {
    expect(snap.cells[1]?.[2]?.value).toBe(8);
    expect(snap.cells[3]?.[2]?.value).toBe(5);
  });
});

describe("parseHtmlTable — defensive handling", () => {
  it("returns null for non-table HTML", () => {
    expect(parseHtmlTable("<p>nothing here</p>")).toBeNull();
    expect(parseHtmlTable("")).toBeNull();
  });

  it("expands colspan into trailing empty cells", () => {
    const snap = parseHtmlTable(
      `<table><tr><td colspan="3">Header</td></tr><tr><td>a</td><td>b</td><td>c</td></tr></table>`
    )!;
    expect(snap.width).toBe(3);
    expect(snap.cells[0]?.[0]?.value).toBe("Header");
    expect(snap.cells[0]?.[1]).toBeNull();
    expect(snap.cells[0]?.[2]).toBeNull();
  });

  it("decodes named and numeric HTML entities", () => {
    const snap = parseHtmlTable(
      `<table><tr><td>caf&eacute;</td><td>a&amp;b</td><td>&#8364;5</td></tr></table>`
    )!;
    // We don't decode `&eacute;` (rare entity); but `&amp;` and
    // numeric entities must round-trip.
    expect(snap.cells[0]?.[1]?.value).toBe("a&b");
    expect(snap.cells[0]?.[2]?.value).toBe("€5");
  });

  it("renders <br> as a real newline inside a cell", () => {
    const snap = parseHtmlTable(`<table><tr><td>line1<br>line2</td></tr></table>`)!;
    expect(snap.cells[0]?.[0]?.value).toBe("line1\nline2");
  });

  it("picks the LARGER table when nested wrapper tables are present", () => {
    // Office HTML often wraps the data table inside a styling outer
    // table. We pick the table with the most cells, not just the
    // first one.
    const html = `<table><tr><td>outer</td></tr></table><table><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>`;
    const snap = parseHtmlTable(html)!;
    expect(snap.height).toBe(2);
    expect(snap.width).toBe(2);
  });
});

describe("parseFingerprintHtml — fingerprint round-trip", () => {
  it("recovers an exact snapshot from a fingerprinted table", () => {
    const original: XlsxClipboardSnapshot = {
      origin: { sheet: "Inventory", range: "A1:B2" },
      width: 2,
      height: 2,
      cells: [
        [{ value: "a", styleId: 4 }, { value: 1 }],
        [{ value: null, formula: "A1*2" }, { value: true }],
      ],
      merges: [{ r0: 0, c0: 1, r1: 1, c1: 1 }],
    };
    const fp = encodeURIComponent(JSON.stringify(original));
    const html = `<table data-xlsx-fingerprint="${fp}"><tr><td>a</td><td>1</td></tr></table>`;
    const recovered = parseFingerprintHtml(html);
    expect(recovered).toEqual(original);
  });

  it("returns null when no fingerprint is present", () => {
    expect(parseFingerprintHtml("<table><tr><td>x</td></tr></table>")).toBeNull();
  });
});

describe("parseExternalClipboard — priority order", () => {
  it("prefers our fingerprint over generic HTML", () => {
    const original: XlsxClipboardSnapshot = {
      origin: { sheet: "S", range: "A1" },
      width: 1,
      height: 1,
      cells: [[{ value: "from-fingerprint" }]],
      merges: [],
    };
    const fp = encodeURIComponent(JSON.stringify(original));
    const html = `<table data-xlsx-fingerprint="${fp}"><tr><td>from-html</td></tr></table>`;
    const out = parseExternalClipboard({ html, text: "from-text" })!;
    expect(out.cells[0]?.[0]?.value).toBe("from-fingerprint");
  });

  it("falls through to generic HTML when the fingerprint is missing", () => {
    const out = parseExternalClipboard({
      html: fixture("excel-desktop.html"),
      text: "ignored",
    })!;
    expect(out.cells[0]?.[0]?.value).toBe("Region");
  });

  it("falls through to plain text when no HTML is provided", () => {
    const out = parseExternalClipboard({ text: "1\t2\n3\t4" })!;
    expect(out.height).toBe(2);
    expect(out.width).toBe(2);
    expect(out.cells[1]?.[1]?.value).toBe(4);
  });

  it("handles a German CSV (semicolon delimiter, quoted fields with commas)", () => {
    const out = parseExternalClipboard({ text: fixture("csv-de.csv") })!;
    expect(out.height).toBe(4);
    expect(out.width).toBe(3);
    expect(out.cells[0]?.[0]?.value).toBe("Produkt");
    expect(out.cells[1]?.[0]?.value).toBe("Schraube, kurz");
    expect(out.cells[2]?.[0]?.value).toBe("Schraube; lang");
    // We don't decimal-localize on import — the German `0,45` lands
    // as a string, not 0.45. Mirrors Excel's import behaviour.
    expect(out.cells[1]?.[2]?.value).toBe("0,45");
  });

  it("handles a multi-line CSV cell quoted with embedded newline", () => {
    const out = parseExternalClipboard({ text: fixture("multiline.csv") })!;
    expect(out.height).toBe(3);
    expect(out.width).toBe(3);
    expect(out.cells[1]?.[1]?.value).toBe("First line\nSecond line");
    expect(out.cells[1]?.[2]?.value).toBe(true);
    expect(out.cells[2]?.[2]?.value).toBe(false);
  });

  it("returns null when nothing is parseable", () => {
    expect(parseExternalClipboard({})).toBeNull();
    expect(parseExternalClipboard({ html: "", text: "" })).toBeNull();
  });
});
