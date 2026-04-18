/**
 * Unit conversions between the canonical points and each format's
 * native font-size unit. Rounding uses `Math.round` so common pt
 * sizes (8, 10, 11, 12, …) round-trip cleanly through any format.
 */

/** OOXML w:sz — half-points. 22 ≙ 11pt. */
export function ptToHalfPoints(pt: number): number {
  return Math.round(pt * 2);
}

export function halfPointsToPt(halfPoints: number): number {
  return halfPoints / 2;
}

/** OOXML a:rPr/@sz — hundredths of a point. 1100 ≙ 11pt. */
export function ptToHundredthsOfPt(pt: number): number {
  return Math.round(pt * 100);
}

export function hundredthsOfPtToPt(hundredths: number): number {
  return hundredths / 100;
}
