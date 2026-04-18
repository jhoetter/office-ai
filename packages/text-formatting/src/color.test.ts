import { describe, expect, it } from "vitest";
import { normalizeColor, renderColor } from "./color";

describe("normalizeColor", () => {
  it("strips '#' and lowercases", () => {
    expect(normalizeColor("#FF8800")).toBe("ff8800");
    expect(normalizeColor("FF8800")).toBe("ff8800");
  });

  it("expands CSS shorthand", () => {
    expect(normalizeColor("f80")).toBe("ff8800");
    expect(normalizeColor("#FFF")).toBe("ffffff");
  });

  it("drops alpha prefix", () => {
    expect(normalizeColor("FFFF8800")).toBe("ff8800");
    expect(normalizeColor("00ff8800")).toBe("ff8800");
  });

  it("returns undefined for invalid input", () => {
    expect(normalizeColor("")).toBeUndefined();
    expect(normalizeColor(null)).toBeUndefined();
    expect(normalizeColor(undefined)).toBeUndefined();
    expect(normalizeColor("not a color")).toBeUndefined();
    expect(normalizeColor("#zzzzzz")).toBeUndefined();
    expect(normalizeColor("12345")).toBeUndefined();
  });
});

describe("renderColor", () => {
  it("prefixes '#'", () => {
    expect(renderColor("ff8800")).toBe("#ff8800");
    expect(renderColor("FF8800")).toBe("#ff8800");
  });
  it("returns undefined for invalid", () => {
    expect(renderColor("")).toBeUndefined();
    expect(renderColor(null)).toBeUndefined();
  });
});
