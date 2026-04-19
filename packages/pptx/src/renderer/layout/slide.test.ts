import { describe, expect, it } from "vitest";
import {
  STAGE_PAD_FRAC_X,
  STAGE_PAD_FRAC_Y,
  slideAspectRatio,
  slideStageViewBox,
  slideViewBox,
  stageAspectRatio,
  SVG_UNIT_PER_EMU,
} from "./slide.js";

const SIZE = { cxEmu: 9_144_000, cyEmu: 6_858_000 } as const;

describe("slideViewBox", () => {
  it("starts at the origin and matches the slide dimensions in user units", () => {
    const cx = SIZE.cxEmu * SVG_UNIT_PER_EMU;
    const cy = SIZE.cyEmu * SVG_UNIT_PER_EMU;
    expect(slideViewBox(SIZE)).toBe(`0 0 ${round2(cx)} ${round2(cy)}`);
  });
});

describe("slideStageViewBox", () => {
  it("centers the slide inside an extended viewport with a negative origin", () => {
    const vb = slideStageViewBox(SIZE).split(" ").map(Number);
    expect(vb).toHaveLength(4);
    const [x, y, w, h] = vb;
    const cx = SIZE.cxEmu * SVG_UNIT_PER_EMU;
    const cy = SIZE.cyEmu * SVG_UNIT_PER_EMU;
    // Negative origin proves shapes positioned in negative EMU still
    // sit inside the visible viewport — the whole point of the
    // scratch canvas around the slide.
    expect(x).toBeLessThan(0);
    expect(y).toBeLessThan(0);
    expect(w).toBeGreaterThan(cx);
    expect(h).toBeGreaterThan(cy);
    // Slide rectangle (0..cx, 0..cy) is centred inside the viewport.
    expect(Math.abs(x + (w - cx) / 2)).toBeLessThan(0.01);
    expect(Math.abs(y + (h - cy) / 2)).toBeLessThan(0.01);
  });

  it("padding fraction widens the viewport by 2x the fraction on each axis", () => {
    const vb = slideStageViewBox(SIZE, 0.5, 0.25).split(" ").map(Number);
    const [, , w, h] = vb;
    const cx = SIZE.cxEmu * SVG_UNIT_PER_EMU;
    const cy = SIZE.cyEmu * SVG_UNIT_PER_EMU;
    expect(w).toBeCloseTo(cx * 2, 1);
    expect(h).toBeCloseTo(cy * 1.5, 1);
  });

  it("default pad fractions match the exported STAGE_PAD_FRAC_* constants", () => {
    const a = slideStageViewBox(SIZE);
    const b = slideStageViewBox(SIZE, STAGE_PAD_FRAC_X, STAGE_PAD_FRAC_Y);
    expect(a).toBe(b);
  });
});

describe("stageAspectRatio", () => {
  it("matches slide aspect ratio when pad fractions are equal on both axes", () => {
    expect(stageAspectRatio(SIZE, 0.2, 0.2)).toBeCloseTo(slideAspectRatio(SIZE), 6);
  });

  it("differs from slide aspect ratio when the pad is asymmetric", () => {
    const stage = stageAspectRatio(SIZE, 0.5, 0);
    const slide = slideAspectRatio(SIZE);
    expect(stage).toBeGreaterThan(slide);
  });
});

function round2(v: number): number {
  return Math.round(v * 100) / 100;
}
