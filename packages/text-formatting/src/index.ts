export {
  MIXED,
  type Mixed,
  type MaybeMixed,
  type TextFormat,
  type ActiveTextFormat,
  type TextFormatCapabilities,
  type TextFormatProvider,
  type UnderlineStyle,
} from "./types";

export { ptToHalfPoints, halfPointsToPt, ptToHundredthsOfPt, hundredthsOfPtToPt } from "./units";

export { normalizeColor, renderColor } from "./color";

export { collapse, isMixed, isOnTruthy, valueOr } from "./mixed";

export {
  FONT_SIZES_PT,
  FONT_FAMILIES,
  COLOR_PALETTE,
  HIGHLIGHT_PALETTE,
  nearestHighlight,
  highlightByDocxName,
  type ColorSwatch,
  type HighlightSwatch,
} from "./presets";

export { wrapFontFamily } from "./font-stack";
