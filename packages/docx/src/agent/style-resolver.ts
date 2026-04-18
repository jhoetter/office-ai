import type {
  DocxSnapshot,
  Paragraph,
  ParagraphProperties,
  Run,
  RunProperties,
  StyleDefinition,
  StylesPart,
  ThemePart,
} from "../model/types.js";

/**
 * Style cascade resolver (P3.1 / W2). See spec/docx/style-cascade.md.
 *
 * Walks `docDefaults.rPrDefault` → the paragraph's `styleId` chain
 * (basedOn) → the paragraph's own `pPr.rPr` → the run's `rPr`. Returns
 * the absolute, fully-merged `RunProperties` / `ParagraphProperties`
 * that the toolbar should display.
 *
 * Pure read-only function. Cycle-safe (depth-capped + visited set).
 */

const MAX_BASED_ON_DEPTH = 16;

/**
 * Word-default font scheme used when a document carries theme refs
 * (`<w:rFonts w:asciiTheme="majorHAnsi"/>`) but ships no
 * `word/theme/theme1.xml` to resolve them — most commonly synthetic
 * test fixtures and hand-written demo docs.
 *
 * Values match what `File → New Document` produces in Word for
 * Microsoft 365 (2024+), which switched the default font scheme from
 * Calibri/Calibri Light to Aptos. Using these values means a doc that
 * round-trips through us *without* shipping a theme part still renders
 * the same way Word does — which is exactly the invariant the user
 * spotted as broken when "Calibri shows in the editor, Aptos shows in
 * Word".
 */
export const WORD_DEFAULT_THEME_FONTS: Readonly<Record<string, string>> = {
  majorHAnsi: "Aptos Display",
  minorHAnsi: "Aptos",
  majorAscii: "Aptos Display",
  minorAscii: "Aptos",
  majorBidi: "Aptos Display",
  minorBidi: "Aptos",
  majorEastAsia: "Aptos Display",
  minorEastAsia: "Aptos",
};

/**
 * Resolve a `<w:rFonts w:asciiTheme="…"/>` reference to a literal
 * typeface. Consults the document's typed {@link ThemePart} first,
 * falls back to {@link WORD_DEFAULT_THEME_FONTS}. Exported so tests
 * (and the toolbar's font dropdown, eventually) can probe the
 * mapping.
 */
export function resolveThemeFont(theme: ThemePart | undefined, themeRef: string): string | undefined {
  if (theme) {
    if (themeRef === "majorHAnsi" || themeRef === "majorAscii" || themeRef === "majorBidi") {
      if (theme.majorFont.latin) return theme.majorFont.latin;
    }
    if (themeRef === "minorHAnsi" || themeRef === "minorAscii" || themeRef === "minorBidi") {
      if (theme.minorFont.latin) return theme.minorFont.latin;
    }
    if (themeRef === "majorEastAsia") return theme.majorFont.ea ?? theme.majorFont.latin;
    if (themeRef === "minorEastAsia") return theme.minorFont.ea ?? theme.minorFont.latin;
  }
  return WORD_DEFAULT_THEME_FONTS[themeRef];
}

export function resolveEffectiveRpr(
  snapshot: DocxSnapshot,
  paragraphIndex: number,
  runIndex?: number
): RunProperties {
  const paragraph = paragraphAt(snapshot, paragraphIndex);
  const styles = snapshot.root.styles;

  let acc: RunProperties = {};
  if (styles?.docDefaults.rPrDefault) acc = mergeRpr(acc, styles.docDefaults.rPrDefault);

  if (paragraph) {
    const styleId = paragraph.properties.styleId;
    if (styleId && styles) {
      acc = mergeRpr(acc, collectStyleRpr(styles, styleId));
    }
    if (runIndex !== undefined) {
      const run = runAt(paragraph, runIndex);
      if (run) acc = mergeRpr(acc, run.properties);
    }
  }

  return projectThemeFont(acc, snapshot.root.theme);
}

export function resolveEffectivePpr(snapshot: DocxSnapshot, paragraphIndex: number): ParagraphProperties {
  const paragraph = paragraphAt(snapshot, paragraphIndex);
  const styles = snapshot.root.styles;

  let acc: ParagraphProperties = {};
  if (styles?.docDefaults.pPrDefault) acc = mergePpr(acc, styles.docDefaults.pPrDefault);

  if (paragraph) {
    const styleId = paragraph.properties.styleId;
    if (styleId && styles) {
      acc = mergePpr(acc, collectStylePpr(styles, styleId));
    }
    acc = mergePpr(acc, paragraph.properties);
  }

  return acc;
}

function paragraphAt(snapshot: DocxSnapshot, paragraphIndex: number): Paragraph | null {
  const block = snapshot.root.body[paragraphIndex];
  if (!block || block.kind !== "paragraph") return null;
  return block;
}

function runAt(paragraph: Paragraph, runIndex: number): Run | null {
  let i = 0;
  for (const child of paragraph.children) {
    if (child.kind === "run") {
      if (i === runIndex) return child;
      i++;
    }
  }
  return null;
}

/**
 * Collapse the basedOn chain root → leaf into a single `RunProperties`.
 * Root applies first, leaf last (so leaf wins on conflicts), matching
 * Word's cascade.
 */
function collectStyleRpr(styles: StylesPart, styleId: string): RunProperties {
  const chain = walkBasedOnChain(styles, styleId);
  let acc: RunProperties = {};
  for (const def of chain) {
    if (def.rPr) acc = mergeRpr(acc, def.rPr);
  }
  return acc;
}

function collectStylePpr(styles: StylesPart, styleId: string): ParagraphProperties {
  const chain = walkBasedOnChain(styles, styleId);
  let acc: ParagraphProperties = {};
  for (const def of chain) {
    if (def.pPr) acc = mergePpr(acc, def.pPr);
  }
  return acc;
}

/**
 * Returns the basedOn chain in root-to-leaf order. A linked style chain
 * `Heading1 → basedOn=Normal → basedOn=undefined` returns
 * `[Normal, Heading1]`. Cycles and missing parents terminate the walk.
 */
function walkBasedOnChain(styles: StylesPart, leafId: string): ReadonlyArray<StyleDefinition> {
  const visited = new Set<string>();
  const reverse: StyleDefinition[] = [];
  let current: string | undefined = leafId;
  let depth = 0;
  while (current && !visited.has(current) && depth < MAX_BASED_ON_DEPTH) {
    visited.add(current);
    const def: StyleDefinition | undefined = styles.styles.get(current);
    if (!def) break;
    reverse.push(def);
    current = def.basedOn;
    depth++;
  }
  reverse.reverse();
  return reverse;
}

/**
 * Merge two `RunProperties`. Right wins on conflicts. `opaqueProps` are
 * concatenated rather than overridden — the cascade resolver doesn't
 * try to dedupe by tag here because the toolbar doesn't read them; a
 * future "render from typed model" path will need a smarter merge.
 */
function mergeRpr(a: RunProperties, b: RunProperties): RunProperties {
  const out: { -readonly [K in keyof RunProperties]: RunProperties[K] } = { ...a };
  if (b.bold !== undefined) out.bold = b.bold;
  if (b.italic !== undefined) out.italic = b.italic;
  if (b.underline !== undefined) out.underline = b.underline;
  if (b.strike !== undefined) out.strike = b.strike;
  // `<w:rFonts>` is treated as a single property by Word's style
  // cascade: if the child level supplies one (literal *or* theme),
  // the parent's `<w:rFonts>` is replaced wholesale, not merged
  // attribute-by-attribute. So a `Heading1` style that only carries
  // `<w:rFonts w:asciiTheme="majorHAnsi"/>` drops the `Calibri` from
  // `docDefaults.rPrDefault`. This matches Word 2024+ rendering of
  // the bundled welcome doc (headings render as Aptos).
  const childHasRFonts =
    b.fontFamily !== undefined ||
    b.fontFamilyAsciiTheme !== undefined ||
    b.fontFamilyHAnsiTheme !== undefined;
  if (childHasRFonts) {
    if (b.fontFamily !== undefined) out.fontFamily = b.fontFamily;
    else delete out.fontFamily;
    if (b.fontFamilyAsciiTheme !== undefined) out.fontFamilyAsciiTheme = b.fontFamilyAsciiTheme;
    else delete out.fontFamilyAsciiTheme;
    if (b.fontFamilyHAnsiTheme !== undefined) out.fontFamilyHAnsiTheme = b.fontFamilyHAnsiTheme;
    else delete out.fontFamilyHAnsiTheme;
  }
  if (b.fontSize !== undefined) out.fontSize = b.fontSize;
  if (b.color !== undefined) out.color = b.color;
  if (b.highlight !== undefined) out.highlight = b.highlight;
  return out;
}

/**
 * Final step of {@link resolveEffectiveRpr}. Mirrors Word's per-attribute
 * resolution rule for `<w:rFonts>`: a literal `w:ascii` always wins
 * over `w:asciiTheme` when both end up on the merged element. So if
 * the cascade produced both — common when a Heading style only
 * carries a theme ref but `docDefaults.rPrDefault` carries a literal
 * font — the literal one is what Word renders.
 *
 * When the literal slot is empty *and* a theme ref is present, we
 * project the theme ref through the loaded {@link ThemePart} (or the
 * Word-default fallback map) into a concrete typeface and write it
 * back into `fontFamily`. The theme attributes themselves are kept on
 * the result so a downstream serializer round-trip stays lossless.
 */
function projectThemeFont(rpr: RunProperties, theme: ThemePart | undefined): RunProperties {
  if (rpr.fontFamily) return rpr;
  const themeRef = rpr.fontFamilyAsciiTheme ?? rpr.fontFamilyHAnsiTheme;
  if (!themeRef) return rpr;
  const projected = resolveThemeFont(theme, themeRef);
  if (!projected) return rpr;
  return { ...rpr, fontFamily: projected };
}

function mergePpr(a: ParagraphProperties, b: ParagraphProperties): ParagraphProperties {
  const out: { -readonly [K in keyof ParagraphProperties]: ParagraphProperties[K] } = { ...a };
  if (b.styleId !== undefined) out.styleId = b.styleId;
  if (b.alignment !== undefined) out.alignment = b.alignment;
  if (b.indentation !== undefined) out.indentation = { ...(a.indentation ?? {}), ...b.indentation };
  if (b.spacing !== undefined) out.spacing = { ...(a.spacing ?? {}), ...b.spacing };
  if (b.numbering !== undefined) out.numbering = b.numbering;
  return out;
}
