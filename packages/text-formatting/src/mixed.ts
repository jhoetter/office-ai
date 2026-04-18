import { MIXED, type MaybeMixed } from "./types";

/**
 * Collapse an iterable of values into a MaybeMixed.
 * - All values strictly equal → that value
 * - All values undefined → undefined
 * - Otherwise → MIXED
 *
 * `eq` defaults to ===. Pass a custom comparator for structurally
 * compared types (e.g. underline variants where `true` and "single"
 * mean the same thing).
 */
export function collapse<T>(
  values: Iterable<T | undefined>,
  eq: (a: T, b: T) => boolean = (a, b) => a === b
): MaybeMixed<T> {
  let first: T | undefined;
  let seen = false;
  for (const v of values) {
    if (!seen) {
      first = v;
      seen = true;
      continue;
    }
    if (first === undefined && v === undefined) continue;
    if (first === undefined || v === undefined) return MIXED;
    if (!eq(first, v)) return MIXED;
  }
  return first;
}

/** True iff a MaybeMixed has the boolean truthy state (false if MIXED or undefined). */
export function isOnTruthy(value: MaybeMixed<boolean>): boolean {
  return value === true;
}

/** Read a MaybeMixed as a definite value, treating MIXED and undefined as fallback. */
export function valueOr<T>(value: MaybeMixed<T>, fallback: T): T {
  if (value === MIXED || value === undefined) return fallback;
  return value;
}

export function isMixed<T>(value: MaybeMixed<T>): boolean {
  return value === MIXED;
}
