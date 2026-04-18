import { ooxml } from "@officeai/core";
import type { ThemeFontEntry, ThemePart } from "../model/types.js";
import { DocxParseError } from "./errors.js";
import { attrOf, elementEntries, findElementEntry } from "./xml-helpers.js";

const DEFAULT_THEME_PART = "word/theme/theme1.xml";

/**
 * Parse `word/theme/theme1.xml` into a typed {@link ThemePart}. The
 * implementation only models `<a:fontScheme>` because that is what the
 * style cascade resolver needs to translate `<w:rFonts w:asciiTheme="…"/>`
 * to a literal typeface ("Aptos Display" in Word 2024+, "Calibri Light"
 * in Word 2013-2023, whatever a custom theme defines).
 *
 * The full DrawingML theme (color scheme, format scheme, custom font
 * overrides, …) is not modeled here — the byte cache in the OPC
 * container preserves it across the round-trip.
 *
 * Returns `undefined` when the package has no theme part (synthetic
 * fixtures, the legacy welcome doc).
 *
 * Discovery: today we look at the canonical path
 * `word/theme/theme1.xml` only. The relationships that point at it
 * (`word/_rels/styles.xml.rels` → `themeRef`) are still parsed and
 * round-tripped via {@link parseRelationshipsParts}; we just don't
 * follow them. Documents in the wild can name the theme part
 * differently, but Word never actually does.
 */
export function parseThemePart(container: ooxml.OoxmlContainer): ThemePart | undefined {
  if (!container.has(DEFAULT_THEME_PART)) return undefined;

  let tree: unknown;
  try {
    tree = ooxml.parseXml(container.readText(DEFAULT_THEME_PART));
  } catch (err) {
    throw new DocxParseError("invalid-xml", "Failed to parse theme1.xml", {
      partPath: DEFAULT_THEME_PART,
      cause: err,
    });
  }
  if (!Array.isArray(tree)) return undefined;

  const root = findElementEntry(tree as unknown[], "a:theme");
  if (!root) return undefined;

  const themeElements = findElementEntry((root["a:theme"] as unknown[] | undefined) ?? [], "a:themeElements");
  if (!themeElements) return undefined;

  const fontScheme = findElementEntry(
    (themeElements["a:themeElements"] as unknown[] | undefined) ?? [],
    "a:fontScheme"
  );
  if (!fontScheme) return undefined;

  const fontSchemeChildren = (fontScheme["a:fontScheme"] as unknown[] | undefined) ?? [];
  const majorFontEntry = findElementEntry(fontSchemeChildren, "a:majorFont");
  const minorFontEntry = findElementEntry(fontSchemeChildren, "a:minorFont");

  // Defensive defaults — Word always ships both, but a hand-edited
  // theme could be missing one. We fall back to the Word 2024 default
  // typefaces so the resolver still gets a usable answer.
  const majorFont = majorFontEntry
    ? parseFontEntry(majorFontEntry, "a:majorFont")
    : { latin: "Aptos Display" };
  const minorFont = minorFontEntry ? parseFontEntry(minorFontEntry, "a:minorFont") : { latin: "Aptos" };

  return {
    partPath: DEFAULT_THEME_PART,
    majorFont,
    minorFont,
  };
}

function parseFontEntry(entry: Record<string, unknown>, tag: "a:majorFont" | "a:minorFont"): ThemeFontEntry {
  const children = (entry[tag] as unknown[] | undefined) ?? [];
  let latin: string | undefined;
  let ea: string | undefined;
  let cs: string | undefined;
  for (const c of elementEntries(children)) {
    const childTag = ooxml.getTag(c);
    if (childTag === "a:latin") latin = attrOf(c, "typeface");
    else if (childTag === "a:ea") ea = attrOf(c, "typeface");
    else if (childTag === "a:cs") cs = attrOf(c, "typeface");
  }
  // Latin is the only required field — the resolver fans out to it
  // for runs that target the Latin script, which is everything the
  // editor authors today.
  const out: { -readonly [K in keyof ThemeFontEntry]: ThemeFontEntry[K] } = { latin: latin ?? "" };
  if (ea) out.ea = ea;
  if (cs) out.cs = cs;
  return out;
}
