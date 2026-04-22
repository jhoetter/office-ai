import { describe, expect, it } from "vitest";
import {
  EMU_PER_CM,
  EMU_PER_INCH,
  TWIPS_PER_CM,
  TWIPS_PER_INCH,
  buildTicks,
  emuToUnit,
  isMajorTick,
  isMetricLocale,
  rulerUnitForLocale,
  twipsToUnit,
} from "./units";

describe("ruler/units", () => {
  describe("isMetricLocale", () => {
    it("treats US/UK/Liberia/Myanmar as imperial", () => {
      expect(isMetricLocale("en-US")).toBe(false);
      expect(isMetricLocale("en-GB")).toBe(false);
      expect(isMetricLocale("en-LR")).toBe(false);
      expect(isMetricLocale("my-MM")).toBe(false);
    });
    it("treats DE / FR / es-MX as metric", () => {
      expect(isMetricLocale("de-DE")).toBe(true);
      expect(isMetricLocale("fr-FR")).toBe(true);
      expect(isMetricLocale("es-MX")).toBe(true);
    });
    it("is case-insensitive", () => {
      expect(isMetricLocale("EN-us")).toBe(false);
      expect(isMetricLocale("DE-de")).toBe(true);
    });
  });

  describe("rulerUnitForLocale", () => {
    it("imperial → in/0.5", () => {
      expect(rulerUnitForLocale("en-US")).toEqual({ unit: "in", step: 0.5 });
    });
    it("metric → cm/1", () => {
      expect(rulerUnitForLocale("de-DE")).toEqual({ unit: "cm", step: 1 });
    });
  });

  describe("buildTicks", () => {
    it("anchors ticks to zero regardless of start", () => {
      // Start at -1.3 cm → first tick is -1, last tick at 5.
      expect(buildTicks(-1.3, 5, 1)).toEqual([-1, 0, 1, 2, 3, 4, 5]);
    });
    it("supports half-inch steps", () => {
      expect(buildTicks(0, 2, 0.5)).toEqual([0, 0.5, 1, 1.5, 2]);
    });
    it("returns [] for non-positive step", () => {
      expect(buildTicks(0, 5, 0)).toEqual([]);
      expect(buildTicks(0, 5, -1)).toEqual([]);
    });
    it("includes endpoints when they fall exactly on a tick", () => {
      expect(buildTicks(0, 3, 1).at(-1)).toBe(3);
    });
  });

  describe("isMajorTick", () => {
    it("marks integers as major", () => {
      expect(isMajorTick(0)).toBe(true);
      expect(isMajorTick(1)).toBe(true);
      expect(isMajorTick(-2)).toBe(true);
    });
    it("marks 0.5 as minor", () => {
      expect(isMajorTick(0.5)).toBe(false);
      expect(isMajorTick(2.5)).toBe(false);
    });
  });

  describe("conversions", () => {
    it("twipsToUnit", () => {
      expect(twipsToUnit(TWIPS_PER_INCH, "in")).toBeCloseTo(1, 6);
      expect(twipsToUnit(TWIPS_PER_CM, "cm")).toBeCloseTo(1, 6);
    });
    it("emuToUnit", () => {
      expect(emuToUnit(EMU_PER_INCH, "in")).toBeCloseTo(1, 6);
      expect(emuToUnit(EMU_PER_CM, "cm")).toBeCloseTo(1, 6);
    });
    it("standard PowerPoint slide (10in × 7.5in) ↔ 25.4cm × 19.05cm", () => {
      const cxEmu = 10 * EMU_PER_INCH;
      const cyEmu = 7.5 * EMU_PER_INCH;
      expect(emuToUnit(cxEmu, "in")).toBeCloseTo(10, 6);
      expect(emuToUnit(cyEmu, "in")).toBeCloseTo(7.5, 6);
      expect(emuToUnit(cxEmu, "cm")).toBeCloseTo(25.4, 4);
      expect(emuToUnit(cyEmu, "cm")).toBeCloseTo(19.05, 4);
    });
  });
});
