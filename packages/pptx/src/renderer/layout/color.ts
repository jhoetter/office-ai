/**
 * Theme-aware color resolution. Today we only handle direct sRGB colors
 * (the common case). Scheme colors fall back to a default palette;
 * P1 follow-up will plumb the actual theme part through the renderer.
 */

export interface ThemeColorScheme {
  readonly accent1: string;
  readonly accent2: string;
  readonly accent3: string;
  readonly accent4: string;
  readonly accent5: string;
  readonly accent6: string;
  readonly tx1: string;
  readonly tx2: string;
  readonly bg1: string;
  readonly bg2: string;
  readonly hlink: string;
  readonly folHlink: string;
}

export const DEFAULT_THEME: ThemeColorScheme = {
  accent1: "5B9BD5",
  accent2: "ED7D31",
  accent3: "A5A5A5",
  accent4: "FFC000",
  accent5: "4472C4",
  accent6: "70AD47",
  tx1: "000000",
  tx2: "44546A",
  bg1: "FFFFFF",
  bg2: "E7E6E6",
  hlink: "0563C1",
  folHlink: "954F72",
};

export type ColorRef =
  | { kind: "srgb"; hex: string }
  | { kind: "sysClr"; hex: string }
  | { kind: "scheme"; name: keyof ThemeColorScheme }
  | { kind: "unsupported" };

export function resolveColor(ref: ColorRef, theme: ThemeColorScheme = DEFAULT_THEME): string {
  switch (ref.kind) {
    case "srgb":
      return ref.hex;
    case "sysClr":
      return ref.hex;
    case "scheme":
      return theme[ref.name];
    case "unsupported":
      return theme.tx1;
  }
}
