/**
 * office-ai Brand Colors
 *
 * Notion-like aesthetic — clean, minimal, with office-ai brand accents.
 * Primary palette is neutral; brand colors used as accents.
 */

export const colors = {
  /* ── Light mode base ── */
  background: "#FFFFFF",
  foreground: "#37352F",
  secondary: "#787774",
  tertiary: "#C3C2C1",
  divider: "#E9E9E7",
  hover: "#F7F7F5",
  surface: "#FBFBFA",
  accent: "#3B82F6",
  accentLight: "#EAF2FE",

  /* ── Dark mode base ── */
  backgroundDark: "#191919",
  foregroundDark: "#E3E2E0",
  secondaryDark: "#9B9A97",
  tertiaryDark: "#5A5A58",
  dividerDark: "#2F2F2F",
  hoverDark: "#252525",
  surfaceDark: "#202020",
  accentDark: "#60A5FA",
  accentLightDark: "#1E2D45",

  /* ── Brand accents (office-ai) ── */
  officeBlue: "#3B82F6",
  officeBlueLight: "#EAF2FE",
  officeBlueMuted: "#3B82F633",
  aiViolet: "#7C3AED",
  aiVioletLight: "#F1ECFE",
  aiVioletMuted: "#7C3AED33",

  /* ── Semantic status ── */
  warning: "#E57A2E",
  error: "#D84B3E",
  info: "#787774",
  success: "#2F7D59",

  /* ── Neutral grays ── */
  gray50: "#FAFAFA",
  gray100: "#F5F5F5",
  gray200: "#E9E9E7",
  gray300: "#C3C2C1",
  gray400: "#9B9A97",
  gray500: "#787774",
  gray600: "#4B5563",
  gray700: "#374151",
  gray800: "#1F2937",
  gray900: "#111827",
} as const;

export type ColorToken = keyof typeof colors;
