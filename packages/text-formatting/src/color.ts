/**
 * Normalise free-form color input into the canonical lowercase RRGGBB
 * (no leading '#'). Returns undefined for invalid / empty input.
 *
 * Accepted inputs:
 * - "#FF8800" / "FF8800"          → "ff8800"
 * - "FFFF8800" (alpha-prefixed)   → "ff8800"  (alpha dropped)
 * - "#f80" / "f80" (CSS shorthand) → "ff8800"
 * - "" / null / undefined         → undefined
 */
export function normalizeColor(input: string | null | undefined): string | undefined {
  if (input == null) return undefined;
  const raw = input.trim().replace(/^#/, "").toLowerCase();
  if (raw.length === 0) return undefined;
  if (!/^[0-9a-f]+$/.test(raw)) return undefined;
  if (raw.length === 3) {
    return raw
      .split("")
      .map((c) => c + c)
      .join("");
  }
  if (raw.length === 6) return raw;
  if (raw.length === 8) return raw.slice(2);
  return undefined;
}

/** Inverse of normalizeColor for inline styles. */
export function renderColor(rrggbb: string | null | undefined): string | undefined {
  const norm = normalizeColor(rrggbb);
  return norm ? `#${norm}` : undefined;
}
