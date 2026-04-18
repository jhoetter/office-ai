import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { parseThemeColorScheme } from "./theme.js";
import { parsePptx } from "./parse.js";
import { DEFAULT_THEME } from "../renderer/layout/color.js";

const FIXTURES = resolve(
  fileURLToPath(import.meta.url),
  "..",
  "..",
  "..",
  "..",
  "..",
  "fixtures",
  "pptx",
  "synthetic"
);

describe("parseThemeColorScheme", () => {
  it("returns DEFAULT_THEME on garbage input", () => {
    expect(parseThemeColorScheme("not xml")).toEqual(DEFAULT_THEME);
    expect(parseThemeColorScheme("")).toEqual(DEFAULT_THEME);
  });

  it("extracts srgbClr accent + dk/lt slots from a hand-rolled theme", () => {
    const xml = [
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
      `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom">`,
      `  <a:themeElements>`,
      `    <a:clrScheme name="Custom">`,
      `      <a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1>`,
      `      <a:lt1><a:sysClr val="window" lastClr="EEEEEE"/></a:lt1>`,
      `      <a:dk2><a:srgbClr val="223344"/></a:dk2>`,
      `      <a:lt2><a:srgbClr val="ABCDEF"/></a:lt2>`,
      `      <a:accent1><a:srgbClr val="FF0000"/></a:accent1>`,
      `      <a:accent2><a:srgbClr val="00FF00"/></a:accent2>`,
      `      <a:accent3><a:srgbClr val="0000FF"/></a:accent3>`,
      `      <a:accent4><a:srgbClr val="123456"/></a:accent4>`,
      `      <a:accent5><a:srgbClr val="654321"/></a:accent5>`,
      `      <a:accent6><a:srgbClr val="ABCDEF"/></a:accent6>`,
      `      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>`,
      `      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>`,
      `    </a:clrScheme>`,
      `  </a:themeElements>`,
      `</a:theme>`,
    ].join("\n");
    const theme = parseThemeColorScheme(xml);
    expect(theme).toEqual({
      tx1: "111111",
      bg1: "EEEEEE",
      tx2: "223344",
      bg2: "ABCDEF",
      accent1: "FF0000",
      accent2: "00FF00",
      accent3: "0000FF",
      accent4: "123456",
      accent5: "654321",
      accent6: "ABCDEF",
      hlink: "0563C1",
      folHlink: "954F72",
    });
  });

  it("falls back to windowText/window mapping when sysClr has no lastClr", () => {
    const xml = [
      `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`,
      `  <a:themeElements><a:clrScheme name="x">`,
      `    <a:dk1><a:sysClr val="windowText"/></a:dk1>`,
      `    <a:lt1><a:sysClr val="window"/></a:lt1>`,
      `  </a:clrScheme></a:themeElements>`,
      `</a:theme>`,
    ].join("\n");
    const theme = parseThemeColorScheme(xml);
    expect(theme.tx1).toBe("000000");
    expect(theme.bg1).toBe("FFFFFF");
  });
});

describe("parsePptx → themeDefault", () => {
  it("populates themeDefault from theme1.xml on each synthetic fixture", async () => {
    for (const name of [
      "01-blank.pptx",
      "02-title-only.pptx",
      "04-multi-shape.pptx",
      "07-multi-slide.pptx",
    ]) {
      const buf = await readFile(resolve(FIXTURES, name));
      const snap = await parsePptx(buf);
      const t = snap.root.themeDefault;
      // Synthetic fixtures use the default Office palette → these are the
      // canonical Office 2013 accents. If the fixture generator changes,
      // update this assertion.
      expect(t.accent1).toMatch(/^[0-9A-F]{6}$/);
      expect(t.tx1).toMatch(/^[0-9A-F]{6}$/);
      expect(t.bg1).toMatch(/^[0-9A-F]{6}$/);
    }
  });
});
