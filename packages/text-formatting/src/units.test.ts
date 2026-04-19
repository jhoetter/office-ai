import { describe, expect, it } from "vitest";
import { halfPointsToPt, hundredthsOfPtToPt, ptToHalfPoints, ptToHundredthsOfPt } from "./units";

describe("units", () => {
  it("round-trips common pt sizes through half-points", () => {
    for (const pt of [8, 9, 10, 10.5, 11, 12, 14, 18, 24, 72]) {
      expect(halfPointsToPt(ptToHalfPoints(pt))).toBeCloseTo(pt);
    }
  });

  it("round-trips common pt sizes through hundredths", () => {
    for (const pt of [8, 9, 10, 11, 11.5, 12, 14, 24, 72]) {
      expect(hundredthsOfPtToPt(ptToHundredthsOfPt(pt))).toBeCloseTo(pt);
    }
  });

  it("ptToHalfPoints rounds to nearest", () => {
    expect(ptToHalfPoints(11)).toBe(22);
    expect(ptToHalfPoints(11.4)).toBe(23);
    expect(ptToHalfPoints(11.6)).toBe(23);
  });
});
