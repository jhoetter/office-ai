/**
 * `wrapFontFamily` — turn a single OOXML font name into a CSS
 * `font-family` value that survives the user not having the font.
 *
 * The XLSX / DOCX / PPTX renderers all pull the run's `font.name`
 * straight out of the workbook / document and emit it verbatim into
 * an inline `font-family:` style. Microsoft fonts (Calibri, Aptos,
 * Cambria, Times New Roman, Arial, …) are NOT installed on macOS or
 * Linux, so a bare `font-family: Calibri` resolves to the user-agent
 * default (typically a Times-style serif) — wildly wrong relative to
 * what the toolbar's font picker claims.
 *
 * Two layers of defence:
 *
 *   1. `apps/web/app/globals.css` declares `@font-face` blocks that
 *      redefine the common Microsoft family names with `local()` →
 *      open-source twin → bundled woff2 fallbacks. That's where the
 *      heavy lifting happens for the curated set of families.
 *   2. This helper appends `system-ui, sans-serif` after the OOXML
 *      name so any family we *haven't* aliased (e.g. "Calisto MT",
 *      "Helvetica Neue LT") still falls through to a sensible system
 *      sans-serif rather than the UA's default serif.
 *
 * The two layers compose cleanly: known families resolve via the
 * `@font-face` chain BEFORE the trailing system fallback ever fires.
 */

const NEEDS_QUOTES = /[\s,"'()]/;

/**
 * Wrap a font name with a generic system fallback chain. Returns
 * `undefined` for null / empty / whitespace-only input so callers can
 * use `if (fam) ...` to skip emitting a `font-family` declaration
 * entirely (matching the pre-existing behaviour at every emission
 * site).
 *
 * Quoting follows CSS `font-family` conventions: identifiers
 * containing whitespace, commas, parentheses, or quote characters are
 * wrapped in double quotes (with embedded `"` escaped). ASCII names
 * without those characters are passed through unquoted so the output
 * is identical to what hand-written CSS would look like.
 */
export function wrapFontFamily(name: string | null | undefined): string | undefined {
  if (!name) return undefined;
  const trimmed = name.trim();
  if (!trimmed) return undefined;
  const quoted = NEEDS_QUOTES.test(trimmed) ? `"${trimmed.replace(/"/g, '\\"')}"` : trimmed;
  return `${quoted}, system-ui, sans-serif`;
}
