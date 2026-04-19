/**
 * Smart-fill series detectors for `xlsx:fill-range`.
 *
 * Excel's drag-the-corner fill handle picks one of a small set of
 * extrapolation strategies:
 *
 *   - numeric arithmetic progression (1,2,3 → 4,5,6; 2,4,6 → 8,10)
 *   - dates (one-day step, derived from any two consecutive points)
 *   - weekday cycle (Mon, Tue → Wed; honours short / long form + case)
 *   - month cycle (Jan, Feb → Mar; honours short / long form + case)
 *   - text + numeric suffix ("Item 1", "Item 2" → "Item 3")
 *   - formula (re-shifts refs against destination, like paste)
 *   - repeat (anything else: cycle the source verbatim)
 *
 * Detectors are tried in priority order. Each returns either `null`
 * (not a match) or a generator that maps an offset (1-based,
 * relative to the END of the source) to the next value.
 *
 * The handler is responsible for mapping `direction = "up" | "left"`
 * by reversing the source before invoking the detector and then
 * walking the offsets in reverse.
 */

import type { Cell, CellValue } from "../model/types.js";

/** A single sample fed into the detector — a value plus optional formula. */
export interface SeriesSample {
  readonly value: CellValue;
  readonly formula?: { readonly text: string } | undefined;
}

export interface SeriesGenerator {
  readonly id: string;
  /** Returns the value at the i-th step past the source's last sample. */
  next(i: number): CellValue;
  /**
   * If the detector wants to emit a formula too (e.g. the formula
   * detector), it returns the shifted formula text via this hook.
   * The handler will pass `(srcRow, srcCol, dstRow, dstCol)` so the
   * detector can re-anchor refs.
   */
  shiftFormula?(i: number, ctx: { readonly dRow: number; readonly dCol: number }): string | null;
}

export type SeriesDetector = (samples: ReadonlyArray<SeriesSample>) => SeriesGenerator | null;

/**
 * Numeric arithmetic progression. Works for a single sample (step 1)
 * or any longer sample that is a true arithmetic series.
 */
export const numericDetector: SeriesDetector = (samples) => {
  if (samples.length === 0) return null;
  const nums: number[] = [];
  for (const s of samples) {
    if (typeof s.value !== "number" || !Number.isFinite(s.value)) return null;
    nums.push(s.value);
  }
  let step = 1;
  if (nums.length === 1) {
    // Excel's exact behaviour with a single numeric sample is to copy.
    // We follow that — use the repeat detector for that case.
    return null;
  }
  step = nums[1]! - nums[0]!;
  for (let i = 2; i < nums.length; i++) {
    const want = nums[i - 1]! + step;
    if (Math.abs(nums[i]! - want) > 1e-9) return null;
  }
  const tail = nums[nums.length - 1]!;
  return {
    id: "numeric",
    next(i: number): CellValue {
      // Round to drop tiny FP drift on integer steps.
      const v = tail + step * i;
      const rounded = Math.round(v);
      return Math.abs(v - rounded) < 1e-9 ? rounded : v;
    },
  };
};

/**
 * Single-numeric copy: when the user grabs ONE numeric cell and
 * drags, Excel repeats the value. We split this from the arithmetic
 * detector because repeat is also the fallback for non-numeric
 * single samples — cleaner to keep them separate so the priority
 * list stays explicit.
 */
export const repeatDetector: SeriesDetector = (samples) => {
  if (samples.length === 0) return null;
  return {
    id: "repeat",
    next(i: number): CellValue {
      // i is 1-based; cycle through the source.
      const idx = (i - 1) % samples.length;
      return samples[idx]!.value;
    },
  };
};

/* ── Weekday + Month cycles ─────────────────────────────────────────────── */

const WEEKDAYS_LONG = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"] as const;
const WEEKDAYS_SHORT = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"] as const;
const MONTHS_LONG = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
] as const;
const MONTHS_SHORT = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
] as const;

type Casing = "lower" | "upper" | "title";

function recase(word: string, casing: Casing): string {
  if (casing === "lower") return word.toLowerCase();
  if (casing === "upper") return word.toUpperCase();
  return word;
}

function detectCycle(
  samples: ReadonlyArray<SeriesSample>,
  cycleLong: readonly string[],
  cycleShort: readonly string[],
  id: string
): SeriesGenerator | null {
  if (samples.length === 0) return null;
  const strs: string[] = [];
  for (const s of samples) {
    if (typeof s.value !== "string") return null;
    strs.push(s.value);
  }
  // Detect form (long / short) and casing from the first sample.
  const first = strs[0]!;
  let cycle: readonly string[];
  if (cycleLong.some((w) => w.toLowerCase() === first.toLowerCase())) {
    cycle = cycleLong;
  } else if (cycleShort.some((w) => w.toLowerCase() === first.toLowerCase())) {
    cycle = cycleShort;
  } else {
    return null;
  }
  let casing: Casing = "title";
  if (first === first.toLowerCase()) casing = "lower";
  else if (first === first.toUpperCase()) casing = "upper";

  const indices: number[] = [];
  for (const s of strs) {
    const i = cycle.findIndex((w) => w.toLowerCase() === s.toLowerCase());
    if (i < 0) return null;
    indices.push(i);
  }
  let step: number;
  if (indices.length === 1) step = 1;
  else step = (indices[1]! - indices[0]! + cycle.length) % cycle.length;
  for (let i = 2; i < indices.length; i++) {
    const want = (indices[i - 1]! + step) % cycle.length;
    if (indices[i] !== want) return null;
  }
  const last = indices[indices.length - 1]!;
  return {
    id,
    next(i: number): CellValue {
      const idx = (((last + step * i) % cycle.length) + cycle.length) % cycle.length;
      return recase(cycle[idx]!, casing);
    },
  };
}

export const weekdayDetector: SeriesDetector = (samples) =>
  detectCycle(samples, WEEKDAYS_LONG, WEEKDAYS_SHORT, "weekday");

export const monthDetector: SeriesDetector = (samples) =>
  detectCycle(samples, MONTHS_LONG, MONTHS_SHORT, "month");

/* ── Text + numeric suffix ─────────────────────────────────────────────── */

const SUFFIX_RE = /^(.*?)(\d+)\s*$/;

/**
 * "Item 1", "Item 2" → "Item 3". Requires the same prefix on every
 * sample and a strictly arithmetic numeric tail.
 */
export const textNumericDetector: SeriesDetector = (samples) => {
  if (samples.length === 0) return null;
  const parts: { prefix: string; n: number; padTo: number }[] = [];
  for (const s of samples) {
    if (typeof s.value !== "string") return null;
    const m = SUFFIX_RE.exec(s.value);
    if (!m) return null;
    const numText = m[2]!;
    const n = Number.parseInt(numText, 10);
    if (!Number.isFinite(n)) return null;
    parts.push({
      prefix: m[1]!,
      n,
      padTo: numText.length > 1 && numText.startsWith("0") ? numText.length : 0,
    });
  }
  const prefix = parts[0]!.prefix;
  for (const p of parts) {
    if (p.prefix !== prefix) return null;
  }
  let step = 1;
  if (parts.length >= 2) {
    step = parts[1]!.n - parts[0]!.n;
    for (let i = 2; i < parts.length; i++) {
      if (parts[i]!.n - parts[i - 1]!.n !== step) return null;
    }
  } else {
    return null; // single sample → fall through to repeat.
  }
  const tail = parts[parts.length - 1]!.n;
  const padTo = parts[parts.length - 1]!.padTo;
  return {
    id: "text-numeric",
    next(i: number): CellValue {
      const n = tail + step * i;
      const numStr = padTo > 0 ? String(n).padStart(padTo, "0") : String(n);
      return `${prefix}${numStr}`;
    },
  };
};

/* ── Date-like (ISO yyyy-mm-dd) ────────────────────────────────────────── */

const ISO_DATE_RE = /^(\d{4})-(\d{2})-(\d{2})$/;

function parseIso(s: string): Date | null {
  const m = ISO_DATE_RE.exec(s);
  if (!m) return null;
  const d = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  if (Number.isNaN(d.getTime())) return null;
  return d;
}

function formatIso(d: Date): string {
  const y = d.getUTCFullYear().toString().padStart(4, "0");
  const m = (d.getUTCMonth() + 1).toString().padStart(2, "0");
  const day = d.getUTCDate().toString().padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/**
 * ISO-formatted dates ("2024-01-31") with a uniform day step. We
 * intentionally don't try to detect month / year stepping yet —
 * Excel's date detector is rich and heuristic; here we cover the
 * 99% case and let the repeat detector pick up the rest.
 */
export const dateDetector: SeriesDetector = (samples) => {
  if (samples.length === 0) return null;
  const dates: Date[] = [];
  for (const s of samples) {
    if (typeof s.value !== "string") return null;
    const d = parseIso(s.value);
    if (!d) return null;
    dates.push(d);
  }
  let stepDays = 1;
  if (dates.length >= 2) {
    stepDays = Math.round((dates[1]!.getTime() - dates[0]!.getTime()) / 86400000);
    for (let i = 2; i < dates.length; i++) {
      const got = Math.round((dates[i]!.getTime() - dates[i - 1]!.getTime()) / 86400000);
      if (got !== stepDays) return null;
    }
  } else {
    return null; // single date → repeat.
  }
  const tail = dates[dates.length - 1]!;
  return {
    id: "date",
    next(i: number): CellValue {
      const d = new Date(tail.getTime() + i * stepDays * 86400000);
      return formatIso(d);
    },
  };
};

/* ── Detector pipeline ─────────────────────────────────────────────────── */

/**
 * Default detector order. The handler walks them in order and
 * returns the first generator that matches. Repeat is the catch-all
 * so the result is always non-null.
 */
export const DEFAULT_DETECTORS: ReadonlyArray<SeriesDetector> = [
  // Specific shapes first; "repeat" must stay last.
  dateDetector,
  weekdayDetector,
  monthDetector,
  textNumericDetector,
  numericDetector,
  repeatDetector,
];

/**
 * Pick a generator for `samples`. Always returns one — falls back to
 * the repeat detector for opaque content (formula-only sources, mixed
 * types, error cells, …).
 */
export function pickSeries(samples: ReadonlyArray<SeriesSample>): SeriesGenerator {
  for (const det of DEFAULT_DETECTORS) {
    const g = det(samples);
    if (g) return g;
  }
  // repeatDetector is in DEFAULT_DETECTORS so we never get here for
  // a non-empty source. Defensive fallback for empty samples.
  return {
    id: "noop",
    next(): CellValue {
      return null;
    },
  };
}

/**
 * Project a row of cells into the series-sample shape. Used by the
 * handler to share the same detection pipeline regardless of the
 * fill direction.
 */
export function cellsToSamples(cells: ReadonlyArray<Cell | undefined>): SeriesSample[] {
  return cells.map((c) => (c ? { value: c.value, formula: c.formula } : { value: null }));
}
