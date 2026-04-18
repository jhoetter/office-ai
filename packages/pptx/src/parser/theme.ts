/**
 * Pure parser for `ppt/theme/themeN.xml` color schemes.
 *
 * The theme part is preserved verbatim as an `OpaquePart` for round-trip
 * fidelity; this helper additionally resolves the first `a:clrScheme`
 * down to a `ThemeColorScheme` value the renderer can consume to draw
 * `<a:schemeClr>` references with the correct palette.
 *
 * Anything we can't resolve falls back to the matching `DEFAULT_THEME`
 * entry — never throws, never blocks parse.
 */

import { ooxml } from "@officeai/core";
import { DEFAULT_THEME, type ThemeColorScheme } from "../renderer/layout/color.js";
import { attrOf, elementEntries, findElementEntry } from "./xml-helpers.js";

export function parseThemeColorScheme(xml: string): ThemeColorScheme {
  let tree: unknown;
  try {
    tree = ooxml.parseXml(xml);
  } catch {
    return DEFAULT_THEME;
  }
  if (!Array.isArray(tree)) return DEFAULT_THEME;

  const themeEl = findElementEntry(tree as unknown[], "a:theme");
  if (!themeEl) return DEFAULT_THEME;

  const themeElements = childByTag(themeEl, "a:themeElements");
  if (!themeElements) return DEFAULT_THEME;

  const clrScheme = childByTag(themeElements, "a:clrScheme");
  if (!clrScheme) return DEFAULT_THEME;

  const slot = (tag: string, fallback: string): string => extractColor(clrScheme, tag) ?? fallback;

  return {
    accent1: slot("a:accent1", DEFAULT_THEME.accent1),
    accent2: slot("a:accent2", DEFAULT_THEME.accent2),
    accent3: slot("a:accent3", DEFAULT_THEME.accent3),
    accent4: slot("a:accent4", DEFAULT_THEME.accent4),
    accent5: slot("a:accent5", DEFAULT_THEME.accent5),
    accent6: slot("a:accent6", DEFAULT_THEME.accent6),
    tx1: slot("a:dk1", DEFAULT_THEME.tx1),
    tx2: slot("a:dk2", DEFAULT_THEME.tx2),
    bg1: slot("a:lt1", DEFAULT_THEME.bg1),
    bg2: slot("a:lt2", DEFAULT_THEME.bg2),
    hlink: slot("a:hlink", DEFAULT_THEME.hlink),
    folHlink: slot("a:folHlink", DEFAULT_THEME.folHlink),
  };
}

function childByTag(parent: Record<string, unknown>, tag: string): Record<string, unknown> | null {
  const own = ooxml.getTag(parent);
  const subtree = parent[own];
  if (!Array.isArray(subtree)) return null;
  return findElementEntry(subtree as unknown[], tag);
}

function extractColor(clrScheme: Record<string, unknown>, slotTag: string): string | null {
  const slot = childByTag(clrScheme, slotTag);
  if (!slot) return null;
  const slotKey = ooxml.getTag(slot);
  const slotKids = slot[slotKey];
  if (!Array.isArray(slotKids)) return null;
  for (const child of elementEntries(slotKids as unknown[])) {
    const tag = ooxml.getTag(child);
    if (tag === "a:srgbClr") {
      const v = attrOf(child, "val");
      if (v) return v.toUpperCase();
    } else if (tag === "a:sysClr") {
      // Prefer @lastClr (resolved), fall back to @val (semantic name —
      // we map "windowText" → 000000 / "window" → FFFFFF only).
      const last = attrOf(child, "lastClr");
      if (last) return last.toUpperCase();
      const val = attrOf(child, "val");
      if (val === "windowText") return "000000";
      if (val === "window") return "FFFFFF";
    }
    // Ignore prstClr/scrgbClr/etc.; let DEFAULT_THEME fill in.
  }
  return null;
}
