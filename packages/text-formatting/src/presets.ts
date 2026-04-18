/**
 * Curated picker presets shared by every editor's TextFormatBar.
 * Single source of truth — replaces the per-product duplicates that
 * used to live in apps/web/app/lib/format-helpers.ts and the XLSX/PPTX
 * toolbars.
 */

/** Common font-size dropdown values, in points. */
export const FONT_SIZES_PT: ReadonlyArray<number> = [
  8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 60, 72,
];

/** Curated font-family dropdown. Adapters may augment with the active value if not present. */
export const FONT_FAMILIES: ReadonlyArray<string> = [
  "Calibri",
  "Cambria",
  "Times New Roman",
  "Arial",
  "Helvetica",
  "Georgia",
  "Verdana",
  "Tahoma",
  "Courier New",
  "Consolas",
];

export interface ColorSwatch {
  /** Display label. */
  readonly name: string;
  /** Lowercase RRGGBB without '#'. */
  readonly hex: string;
}

/** Common color picker palette (font color). Lowercase RRGGBB. */
export const COLOR_PALETTE: ReadonlyArray<ColorSwatch> = [
  { name: "Default", hex: "000000" },
  { name: "Gray", hex: "595959" },
  { name: "Red", hex: "c00000" },
  { name: "Orange", hex: "ed7d31" },
  { name: "Yellow", hex: "ffc000" },
  { name: "Green", hex: "70ad47" },
  { name: "Blue", hex: "2e75b6" },
  { name: "Purple", hex: "7030a0" },
];

export interface HighlightSwatch {
  /** Display label. */
  readonly name: string;
  /** Lowercase RRGGBB without '#'. */
  readonly hex: string;
  /**
   * For DOCX: the corresponding `<w:highlight w:val="…">` enum value.
   * Other adapters ignore this.
   */
  readonly docxName: string;
}

/**
 * Highlight palette. Hex values approximate Word's swatches and
 * also enable XLSX/PPTX adapters to render the same swatch via
 * cell fill / a:highlight respectively. The DOCX adapter quantises
 * incoming RGB onto the closest entry's `docxName`.
 */
export const HIGHLIGHT_PALETTE: ReadonlyArray<HighlightSwatch> = [
  { name: "Yellow", hex: "ffff00", docxName: "yellow" },
  { name: "Green", hex: "00ff00", docxName: "green" },
  { name: "Cyan", hex: "00ffff", docxName: "cyan" },
  { name: "Magenta", hex: "ff00ff", docxName: "magenta" },
  { name: "Red", hex: "ff0000", docxName: "red" },
  { name: "Olive", hex: "808000", docxName: "darkYellow" },
  { name: "Light gray", hex: "c0c0c0", docxName: "lightGray" },
];

/**
 * Find the closest highlight swatch to an arbitrary RRGGBB. Used by
 * the DOCX adapter to map a free-form color picker output onto the
 * OOXML enum. Distance is squared-RGB.
 */
export function nearestHighlight(rrggbb: string): HighlightSwatch {
  const norm = rrggbb.toLowerCase();
  const r = parseInt(norm.slice(0, 2), 16);
  const g = parseInt(norm.slice(2, 4), 16);
  const b = parseInt(norm.slice(4, 6), 16);
  let best = HIGHLIGHT_PALETTE[0];
  let bestDist = Infinity;
  for (const swatch of HIGHLIGHT_PALETTE) {
    const sr = parseInt(swatch.hex.slice(0, 2), 16);
    const sg = parseInt(swatch.hex.slice(2, 4), 16);
    const sb = parseInt(swatch.hex.slice(4, 6), 16);
    const d = (r - sr) ** 2 + (g - sg) ** 2 + (b - sb) ** 2;
    if (d < bestDist) {
      bestDist = d;
      best = swatch;
    }
  }
  return best;
}

/** Look up a swatch by its DOCX enum name. */
export function highlightByDocxName(name: string): HighlightSwatch | undefined {
  return HIGHLIGHT_PALETTE.find((s) => s.docxName === name);
}
