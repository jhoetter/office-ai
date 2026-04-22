import { describe, expect, it } from "vitest";
import { wrapFontFamily } from "./font-stack";

describe("wrapFontFamily", () => {
  it("returns undefined for null / undefined / empty / whitespace", () => {
    expect(wrapFontFamily(null)).toBeUndefined();
    expect(wrapFontFamily(undefined)).toBeUndefined();
    expect(wrapFontFamily("")).toBeUndefined();
    expect(wrapFontFamily("   ")).toBeUndefined();
    expect(wrapFontFamily("\t\n")).toBeUndefined();
  });

  it("passes through ASCII single-word identifiers unquoted", () => {
    // CSS allows `font-family: Calibri` without quotes for identifiers
    // that don't contain whitespace / commas / parens / quotes.
    expect(wrapFontFamily("Calibri")).toBe("Calibri, system-ui, sans-serif");
    expect(wrapFontFamily("Arial")).toBe("Arial, system-ui, sans-serif");
    expect(wrapFontFamily("Verdana")).toBe("Verdana, system-ui, sans-serif");
  });

  it("quotes multi-word names so the family parses as a single token", () => {
    expect(wrapFontFamily("Times New Roman")).toBe('"Times New Roman", system-ui, sans-serif');
    expect(wrapFontFamily("Courier New")).toBe('"Courier New", system-ui, sans-serif');
    expect(wrapFontFamily("Calibri Light")).toBe('"Calibri Light", system-ui, sans-serif');
    expect(wrapFontFamily("Aptos Display")).toBe('"Aptos Display", system-ui, sans-serif');
  });

  it("trims surrounding whitespace before deciding whether to quote", () => {
    expect(wrapFontFamily("  Calibri  ")).toBe("Calibri, system-ui, sans-serif");
    expect(wrapFontFamily("\tArial\n")).toBe("Arial, system-ui, sans-serif");
  });

  it("escapes embedded double quotes in pathological font names", () => {
    expect(wrapFontFamily('Bad"Name')).toBe('"Bad\\"Name", system-ui, sans-serif');
  });

  it("quotes names with commas, parens, or quote characters", () => {
    expect(wrapFontFamily("Foo,Bar")).toBe('"Foo,Bar", system-ui, sans-serif');
    expect(wrapFontFamily("Foo (Bold)")).toBe('"Foo (Bold)", system-ui, sans-serif');
    expect(wrapFontFamily("O'Reilly")).toBe(`"O'Reilly", system-ui, sans-serif`);
  });
});
