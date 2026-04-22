import { describe, expect, it } from "vitest";
import { slideGridSvgString } from "./SlideGridOverlay.js";

const SLIDE_10x7_5_IN = { cxEmu: 9144000, cyEmu: 6858000 };

describe("slideGridSvgString", () => {
  it("emits the data-testid hook the e2e and view-tab tests rely on", () => {
    const svg = slideGridSvgString(SLIDE_10x7_5_IN, "in");
    expect(svg).toContain('data-testid="pptx-grid-overlay"');
  });

  it("imperial unit → 0.5in (48px@96dpi) major step, 0.125in (12px) minor step", () => {
    const svg = slideGridSvgString(SLIDE_10x7_5_IN, "in");
    // 0.5in = 0.5 * 914400 / 9525 = 48 user units.
    expect(svg).toContain('width="48"');
    expect(svg).toContain('height="48"');
    // Minor step is a quarter of the major: 12 user units.
    expect(svg).toContain('width="12"');
    expect(svg).toContain('height="12"');
  });

  it("metric unit → 1cm (≈ 37.8px@96dpi) major step", () => {
    const svg = slideGridSvgString(SLIDE_10x7_5_IN, "cm");
    // 1 cm = 360000 / 9525 = 37.7952… user units.
    expect(svg).toMatch(/width="37\.7952\d*"/);
    expect(svg).toMatch(/height="37\.7952\d*"/);
  });

  it("background rect spans the entire slide rectangle in user units", () => {
    const svg = slideGridSvgString(SLIDE_10x7_5_IN, "in");
    // 10 in = 9144000 / 9525 = 960 user units; 7.5 in = 720 user units.
    expect(svg).toContain('width="960"');
    expect(svg).toContain('height="720"');
  });

  it("uses userSpaceOnUse so the pattern aligns with slide coords (not pattern bbox)", () => {
    const svg = slideGridSvgString(SLIDE_10x7_5_IN, "in");
    expect(svg).toContain('patternUnits="userSpaceOnUse"');
  });
});
