import { describe, expect, it } from "vitest";
import { selectEngine } from "./select-engine.js";

describe("selectEngine", () => {
  it("defaults to pdfjs when no hints fire", () => {
    expect(selectEngine()).toBe("pdfjs");
    expect(selectEngine({})).toBe("pdfjs");
    expect(selectEngine({ linearized: true })).toBe("pdfjs");
  });

  it("opts into pdfium when the user explicitly prefers fidelity", () => {
    expect(selectEngine({ userPrefersFidelity: true })).toBe("pdfium");
  });

  it("opts into pdfium for known-bad hints", () => {
    expect(selectEngine({ hasUncommonColorSpace: true })).toBe("pdfium");
    expect(selectEngine({ hasType3Fonts: true })).toBe("pdfium");
    expect(selectEngine({ hasCustomCMap: true })).toBe("pdfium");
    expect(selectEngine({ inPdfiumAllowlist: true })).toBe("pdfium");
  });

  it("user preference wins over disabling hints", () => {
    expect(
      selectEngine({
        userPrefersFidelity: true,
        hasUncommonColorSpace: false,
        hasType3Fonts: false,
        hasCustomCMap: false,
      }),
    ).toBe("pdfium");
  });
});
