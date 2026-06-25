import { colors } from "./colors";
import { fontFamily } from "./typography";
import { borderRadius, borderWidth, maxWidth, spacing } from "./spacing";

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

        sl: {
          bg: "var(--sl-bg)",
          surface: "var(--sl-surface)",
          "surface-2": "var(--sl-surface-2)",
          hover: "var(--sl-hover)",
          selected: "var(--sl-sel)",
          line: "var(--sl-line)",
          ink: "var(--sl-ink)",
          muted: "var(--sl-muted)",
          faint: "var(--sl-faint)",
          accent: "var(--sl-accent)",
          "accent-weak": "var(--sl-accent-weak)",
          "accent-ink": "var(--sl-accent-ink)",
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
      spacing: { ...spacing },
      borderRadius: { ...borderRadius },
      borderWidth: { ...borderWidth },
      maxWidth: { ...maxWidth },
    },
  },
};
