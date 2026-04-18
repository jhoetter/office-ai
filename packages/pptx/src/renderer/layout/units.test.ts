import { describe, expect, it } from "vitest";
import {
  DEFAULT_DPI,
  EMU_PER_INCH,
  EMU_PER_PX_AT_96DPI,
  emuToPx,
  fontSizeHundredthsToPx,
  pxToEmu,
} from "./units.js";

describe("EMU/pixel conversions", () => {
  it("constants match the OOXML spec", () => {
    expect(EMU_PER_INCH).toBe(914400);
    expect(DEFAULT_DPI).toBe(96);
    expect(EMU_PER_PX_AT_96DPI).toBe(914400 / 96);
  });

  it("pxToEmu(emuToPx(e)) ≈ e for typical values", () => {
    for (const emu of [0, 9525, 914400, 12345600, 9144000]) {
      expect(pxToEmu(emuToPx(emu))).toBe(emu);
    }
  });

  it("hundredths-of-a-point → px at 96 DPI", () => {
    // 1800 hundredths = 18pt → 24px at 96 DPI
    expect(fontSizeHundredthsToPx(1800)).toBeCloseTo(24, 5);
  });
});
