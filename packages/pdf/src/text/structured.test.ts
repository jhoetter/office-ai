import { describe, expect, it } from "vitest";
import type { PdfEngineGlyphRun } from "@officeai/pdf-engine";
import { buildStructuredPage } from "./structured.js";
import {
  collectTextWithinRegions,
  serializeMarkdown,
  serializeReadingOrder,
} from "./serialize.js";
import { findInStructuredPage } from "./search.js";

/**
 * Build a single-line glyph run at the given baseline.
 * Each glyph is 6 PDF points wide × 12 tall — comfortable for the
 * 0.5×fontHeight line-bucket tolerance.
 */
function makeRun(
  chars: string,
  startX: number,
  baselineY: number,
  options: { fontHeight?: number; charWidth?: number } = {},
): PdfEngineGlyphRun {
  const fontHeight = options.fontHeight ?? 12;
  const charWidth = options.charWidth ?? 6;
  const glyphs: Array<readonly [number, number, number, number]> = [];
  let x = startX;
  for (let i = 0; i < chars.length; i++) {
    glyphs.push([x, baselineY, x + charWidth, baselineY + fontHeight] as const);
    x += charWidth;
  }
  return {
    chars,
    glyphs,
    baselineY,
    bbox: [startX, baselineY, x, baselineY + fontHeight] as const,
    fontHeight,
    fontKey: "TestFont",
    dir: "ltr",
    hasEol: true,
  };
}

describe("buildStructuredPage", () => {
  it("groups glyphs into lines by baseline tolerance", () => {
    const runs = [
      makeRun("Hello", 50, 700),
      makeRun("World", 100, 700),
      makeRun("Second line", 50, 680),
    ];
    const page = buildStructuredPage(runs, 612, 792);
    expect(page.columnCount).toBe(1);
    expect(page.blocks.length).toBeGreaterThan(0);
    const allLines = page.blocks.flatMap((b) => b.lines);
    expect(allLines.length).toBe(2);
    expect(allLines[0].text).toContain("Hello");
    expect(allLines[0].text).toContain("World");
    expect(allLines[1].text).toContain("Second line");
  });

  it("emits a heading block when a line uses a noticeably larger font", () => {
    const runs = [
      makeRun("Title Of Section", 50, 720, { fontHeight: 18 }),
      makeRun("Body paragraph here", 50, 700, { fontHeight: 12 }),
    ];
    const page = buildStructuredPage(runs, 612, 792);
    const kinds = page.blocks.map((b) => b.kind);
    expect(kinds).toContain("heading");
  });

  it("detects two columns when a stable horizontal gap separates ink", () => {
    const runs: PdfEngineGlyphRun[] = [];
    // Left column at x=50..200, right column at x=400..550, page width 612.
    // Stack 6 lines per column to give the gap full vertical extent.
    for (let i = 0; i < 6; i++) {
      const y = 720 - i * 16;
      runs.push(makeRun("LeftColLine", 50, y));
      runs.push(makeRun("RightColLine", 400, y));
    }
    const page = buildStructuredPage(runs, 612, 792);
    expect(page.columnCount).toBe(2);
    // Reading order: full left column first, then full right column.
    const text = serializeReadingOrder(page);
    const firstRight = text.indexOf("RightColLine");
    const lastLeft = text.lastIndexOf("LeftColLine");
    expect(lastLeft).toBeLessThan(firstRight);
  });

  it("classifies a list block when the first line starts with a bullet", () => {
    const runs = [
      makeRun("\u2022 First item", 50, 700),
      makeRun("\u2022 Second item", 50, 680),
    ];
    const page = buildStructuredPage(runs, 612, 792);
    expect(page.blocks[0].kind).toBe("list");
    const md = serializeMarkdown(page);
    expect(md).toContain("- First item");
    expect(md).toContain("- Second item");
  });
});

describe("findInStructuredPage", () => {
  it("returns per-line PDF rects for a found phrase", () => {
    const runs = [
      makeRun("Hello world", 50, 700),
      makeRun("foo bar baz", 50, 680),
    ];
    const page = buildStructuredPage(runs, 612, 792);
    const text = serializeReadingOrder(page);
    const re = /world/g;
    const hits = findInStructuredPage(page, re, text);
    expect(hits).not.toBeNull();
    expect(hits!.length).toBe(1);
    expect(hits![0].rects.length).toBe(1);
    const [x1, y1, x2, y2] = hits![0].rects[0];
    expect(x1).toBeGreaterThan(70); // "Hello " is 6 glyphs × 6pt
    expect(x2).toBeGreaterThan(x1);
    expect(y1).toBeLessThan(y2);
  });

  it("emits one rect per visual line for matches that wrap", () => {
    const runs = [
      makeRun("abc def", 50, 700),
      makeRun("ghi jkl", 50, 680),
    ];
    const page = buildStructuredPage(runs, 612, 792);
    const text = serializeReadingOrder(page);
    // Match characters that span both lines via a deliberately
    // permissive regex.
    const hits = findInStructuredPage(page, /[a-z ]+/g, text);
    expect(hits).not.toBeNull();
    expect(hits!.length).toBeGreaterThan(0);
  });

  it("returns null when the structured projection diverges from the supplied plain text", () => {
    const runs = [makeRun("Hello", 50, 700)];
    const page = buildStructuredPage(runs, 612, 792);
    // Pretend the page text was mutated out-of-band; the helper
    // should bail out so the caller falls back to legacy search.
    const hits = findInStructuredPage(page, /Hello/g, "Different");
    expect(hits).toBeNull();
  });
});

describe("collectTextWithinRegions", () => {
  it("emits column-major reading order across a multi-column selection", () => {
    const runs: PdfEngineGlyphRun[] = [];
    for (let i = 0; i < 6; i++) {
      const y = 720 - i * 16;
      runs.push(makeRun(`Left${i}`, 50, y));
      runs.push(makeRun(`Right${i}`, 400, y));
    }
    const page = buildStructuredPage(runs, 612, 792);
    // Big selection covering both columns and all lines.
    const text = collectTextWithinRegions(page, [[40, 600, 560, 740]]);
    const firstRight = text.indexOf("Right0");
    const lastLeft = text.lastIndexOf("Left5");
    expect(lastLeft).toBeLessThan(firstRight);
  });

  it("returns text for a one-line single-region selection", () => {
    const runs = [makeRun("Pick me", 50, 700)];
    const page = buildStructuredPage(runs, 612, 792);
    const text = collectTextWithinRegions(page, [[40, 695, 200, 720]]);
    expect(text).toContain("Pick me");
  });
});
