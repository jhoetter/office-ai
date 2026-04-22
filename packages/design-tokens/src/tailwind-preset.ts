import { colors } from "./colors";
import { fontFamily } from "./typography";
import { borderRadius, borderWidth } from "./spacing";

/**
 * office-ai Tailwind CSS Preset
 *
 * Intended for use with Tailwind v3 configs. With Tailwind v4 (used in
 * apps/web), tokens are emitted via CSS variables in `globals.css`.
 */
export const officeAiPreset = {
  theme: {
    extend: {
      colors: {
        background: "var(--background)",
        foreground: "var(--foreground)",
        secondary: "var(--secondary)",
        tertiary: "var(--tertiary)",
        divider: "var(--divider)",
        hover: "var(--hover)",
        surface: "var(--surface)",
        accent: "var(--accent)",
        "accent-light": "var(--accent-light)",

        "office-blue": {
          DEFAULT: colors.officeBlue,
          light: colors.officeBlueLight,
          muted: colors.officeBlueMuted,
        },
        "ai-violet": {
          DEFAULT: colors.aiViolet,
          light: colors.aiVioletLight,
          muted: colors.aiVioletMuted,
        },

        warning: colors.warning,
        error: colors.error,
        info: colors.info,
        success: colors.success,
      },
      fontFamily: {
        sans: fontFamily.sans,
        mono: fontFamily.mono,
      },
      borderRadius: { ...borderRadius },
      borderWidth: { ...borderWidth },
      maxWidth: {
        content: "1200px",
        prose: "65ch",
      },
    },
  },
};
