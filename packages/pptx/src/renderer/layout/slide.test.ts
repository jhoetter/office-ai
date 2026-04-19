import { describe, expect, it } from "vitest";
import {
  STAGE_PAD_FRAC_X,
  STAGE_PAD_FRAC_Y,
  computeStageLayout,
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

describe("computeStageLayout", () => {
  it("returns null for a zero-sized stage so callers can no-op", () => {
    expect(computeStageLayout(SIZE, 0, 100)).toBeNull();
    expect(computeStageLayout(SIZE, 100, 0)).toBeNull();
  });

  it("centers the slide card inside a wider-than-tall stage with horizontal letterbox", () => {
    // Stage is 16:9 + extra horizontal margin; slide stays 4:3 (≈ 1.333).
    const stage = { w: 1600, h: 800 };
    const layout = computeStageLayout(SIZE, stage.w, stage.h);
    expect(layout).not.toBeNull();
    const aspect = slideAspectRatio(SIZE);
    expect(layout!.slidePxW / layout!.slidePxH).toBeCloseTo(aspect, 5);
    expect(layout!.slidePxH).toBeCloseTo(stage.h, 5);
    expect(layout!.slidePxLeft).toBeGreaterThan(0);
    expect(layout!.slidePxTop).toBeCloseTo(0, 5);
    expect(layout!.slidePxLeft * 2 + layout!.slidePxW).toBeCloseTo(stage.w, 5);
  });

  it("centers the slide card inside a taller-than-wide stage with vertical letterbox", () => {
    const stage = { w: 800, h: 900 };
    const layout = computeStageLayout(SIZE, stage.w, stage.h);
    expect(layout).not.toBeNull();
    expect(layout!.slidePxW).toBeCloseTo(stage.w, 5);
    expect(layout!.slidePxTop).toBeGreaterThan(0);
    expect(layout!.slidePxLeft).toBeCloseTo(0, 5);
    expect(layout!.slidePxTop * 2 + layout!.slidePxH).toBeCloseTo(stage.h, 5);
  });

  it("scales the slide card by the zoom multiplier", () => {
    const base = computeStageLayout(SIZE, 1200, 900, 1)!;
    const zoomed = computeStageLayout(SIZE, 1200, 900, 1.5)!;
    expect(zoomed.slidePxW).toBeCloseTo(base.slidePxW * 1.5, 5);
    expect(zoomed.slidePxH).toBeCloseTo(base.slidePxH * 1.5, 5);
  });

  it("emits a viewBox where SVG (0,0) lines up with the slide card and (slideW, slideH) lines up with its bottom-right", () => {
    const stage = { w: 1600, h: 900 };
    const layout = computeStageLayout(SIZE, stage.w, stage.h)!;
    const [vbX, vbY, vbW, vbH] = layout.stageViewBox.split(" ").map(Number);
    const slideWUser = SIZE.cxEmu * SVG_UNIT_PER_EMU;
    const slideHUser = SIZE.cyEmu * SVG_UNIT_PER_EMU;
    const userPerPx = vbW / stage.w;
    expect(vbH / stage.h).toBeCloseTo(userPerPx, 5);
    expect(-vbX / userPerPx).toBeCloseTo(layout.slidePxLeft, 1);
    expect(-vbY / userPerPx).toBeCloseTo(layout.slidePxTop, 1);
    expect(slideWUser / userPerPx).toBeCloseTo(layout.slidePxW, 1);
    expect(slideHUser / userPerPx).toBeCloseTo(layout.slidePxH, 1);
  });

  it("yields a negative viewBox origin so off-slide shapes remain visible in the scratch margin", () => {
    // Wider-than-slide stage: viewBox.x must be negative so EMU(<0, 0)
    // shapes (left scratch margin) sit inside the visible viewport.
    const wide = computeStageLayout(SIZE, 1600, 900)!;
    expect(Number(wide.stageViewBox.split(" ")[0])).toBeLessThan(0);
    // Taller-than-slide stage: viewBox.y must be negative for the same
    // reason on the top scratch margin.
    const tall = computeStageLayout(SIZE, 800, 900)!;
    expect(Number(tall.stageViewBox.split(" ")[1])).toBeLessThan(0);
    // Zoom < 1 leaves slack on both axes so both origins go negative.
    const zoomed = computeStageLayout(SIZE, 1200, 900, 0.5)!;
    expect(Number(zoomed.stageViewBox.split(" ")[0])).toBeLessThan(0);
    expect(Number(zoomed.stageViewBox.split(" ")[1])).toBeLessThan(0);
  });
});

function round2(v: number): number {
  return Math.round(v * 100) / 100;
}
