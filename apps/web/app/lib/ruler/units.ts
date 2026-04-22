/**
 * Shared linear-ruler helpers used by the DOCX `PageRuler` (top of the
 * page) and the PPTX `SlideRulers` (top + left of the slide card).
 *
 * Both rulers paint the same visual: ticks at every minor step, with a
 * numeric label only at every major step. The unit and step are picked
 * from the user's locale: imperial (in / 0.5 step) for US/UK/LR/MM,
 * metric (cm / 1 step) elsewhere — matching what Word and PowerPoint do.
 *
 * Pure module (no DOM, no React) so it's trivially testable.
 */

export const TWIPS_PER_INCH = 1440;
export const TWIPS_PER_CM = 567; // 1440 / 2.54

export const EMU_PER_INCH = 914400;
export const EMU_PER_CM = 360000; // 914400 / 2.54

export type RulerUnit = "in" | "cm";

export interface RulerUnitInfo {
  readonly unit: RulerUnit;
  /** Minor tick step in `unit`. */
  readonly step: number;
}

/**
 * Locale-driven unit choice. Mirrors {@link isMetricLocale} from the
 * original DOCX `PageRuler` so the two rulers always agree.
 */
export function rulerUnitForLocale(locale?: string): RulerUnitInfo {
  return isMetricLocale(locale) ? { unit: "cm", step: 1 } : { unit: "in", step: 0.5 };
}

export function isMetricLocale(locale?: string): boolean {
  const raw =
    locale ?? (typeof navigator !== "undefined" ? navigator.language : "en-US") ?? "en-US";
  const lang = raw.toLowerCase();
  if (lang.startsWith("en-us")) return false;
  if (lang.startsWith("en-gb")) return false;
  if (lang.startsWith("en-lr")) return false;
  if (lang === "my-mm") return false;
  return true;
}

/**
 * Build an inclusive list of tick positions from `start` to `end`
 * (in the chosen `unit`) at `step` increments. `start` is typically
 * `0`; PPTX's vertical/horizontal ruler passes a negative value to
 * cover the scratch margin.
 */
export function buildTicks(start: number, end: number, step: number): number[] {
  if (!(step > 0)) return [];
  const ticks: number[] = [];
  // Anchor to zero so positive and negative ticks line up regardless
  // of where `start` falls.
  const firstK = Math.ceil(start / step - 1e-6);
  const lastK = Math.floor(end / step + 1e-6);
  for (let k = firstK; k <= lastK; k++) {
    // `k === 0` short-circuit avoids `-0` when the previous tick was
    // negative — it never affects rendering but trips deep-equality
    // comparisons in tests.
    ticks.push(k === 0 ? 0 : k * step);
  }
  return ticks;
}

/**
 * A tick is "major" when its value (in `unit`) rounds to a whole
 * integer. Major ticks get a longer line and a numeric label.
 */
export function isMajorTick(value: number): boolean {
  return Math.abs(value - Math.round(value)) < 1e-6;
}

/** Convert twips → the chosen ruler unit. */
export function twipsToUnit(twips: number, unit: RulerUnit): number {
  return unit === "cm" ? twips / TWIPS_PER_CM : twips / TWIPS_PER_INCH;
}

/** Convert EMU → the chosen ruler unit. */
export function emuToUnit(emu: number, unit: RulerUnit): number {
  return unit === "cm" ? emu / EMU_PER_CM : emu / EMU_PER_INCH;
}
