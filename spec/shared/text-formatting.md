# Shared Text Formatting

> Format-agnostic vocabulary for run-level text formatting (bold, italic,
> underline, strike, font family, font size, color, highlight) consumed
> by every editor in this monorepo. Each format plugin (DOCX, XLSX,
> PPTX) translates between this normalised shape and its own native
> model.

## Goals

1. **One picker, three products.** The same `FontFamilyPicker`,
   `FontSizePicker`, `ColorPicker`, `HighlightPicker`, and B/I/U/S
   toggles render in the DOCX, XLSX, and PPTX toolbars and behave
   identically.
2. **One canonical unit.** Font sizes are normalised to **points**
   (float). Colors are normalised to lowercase 6-character hex
   without a leading `#`. The shared layer never sees half-points,
   hundredths, or `RRGGBBAA`.
3. **Mixed-state aware.** A toolbar with a heterogeneous selection
   shows `—` for divergent attributes; clicking still applies the
   click value. The shared `MaybeMixed<T>` type encodes this.
4. **Capability-driven UI.** Products that can't express a control
   (e.g. native character highlight in PPTX classic, or font family
   in XLSX as `font.name`) declare it via the `capabilities` flag and
   the UI either hides the control, falls back, or no-ops.

## Canonical types

Defined in `@officeai/text-formatting`:

```typescript
export type UnderlineStyle =
  | "single"
  | "double"
  | "thick"
  | "dotted"
  | "dashed"
  | "wavy";

export interface TextFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | UnderlineStyle;
  strike?: boolean;
  fontFamily?: string;
  /** Canonical font size in points. */
  fontSizePt?: number;
  /** RRGGBB lowercase, no '#'. */
  color?: string;
  /**
   * RRGGBB lowercase, no '#'. DOCX maps to its w:highlight enum
   * via the closest named swatch (see HIGHLIGHT_PALETTE); XLSX maps
   * to cell fill; PPTX writes <a:highlight><a:srgbClr/>.
   */
  highlight?: string;
}

export const MIXED: unique symbol;
export type MaybeMixed<T> = T | typeof MIXED | undefined;

export interface ActiveTextFormat {
  bold: MaybeMixed<boolean>;
  italic: MaybeMixed<boolean>;
  underline: MaybeMixed<boolean | UnderlineStyle>;
  strike: MaybeMixed<boolean>;
  fontFamily: MaybeMixed<string>;
  fontSizePt: MaybeMixed<number>;
  color: MaybeMixed<string>;
  highlight: MaybeMixed<string>;
}
```

The provider contract each editor implements:

```typescript
export interface TextFormatCapabilities {
  /** Native = round-trips as a real character highlight. */
  highlight: "native" | "fill-fallback" | "none";
  underlineVariants: boolean;
  fontFamily: boolean;
  fontSize: boolean;
}

export interface TextFormatProvider {
  /** Cheap, called on every render. */
  getActive(): ActiveTextFormat;
  /** Dispatches the format-specific command via the agent. */
  apply(patch: TextFormat): void;
  hasSelection(): boolean;
  capabilities: TextFormatCapabilities;
}
```

## Unit conversions

| Format | Native unit | Conversion |
|---|---|---|
| DOCX | half-points (`<w:sz w:val="22"/>` = 11 pt) | `pt = halfPoints / 2` |
| XLSX | points (`<sz val="11"/>`) | identity |
| PPTX | hundredths of a point (`<a:rPr sz="1100"/>` = 11 pt) | `pt = hundredths / 100` |

Helpers are exported as `ptToHalfPoints`, `halfPointsToPt`,
`ptToHundredthsOfPt`, `hundredthsOfPtToPt`. Rounding uses
`Math.round` so that 11 pt round-trips through any unit cleanly.

## Color normalisation

| Input | Output |
|---|---|
| `"#FF8800"` | `"ff8800"` |
| `"FF8800"` | `"ff8800"` |
| `"FFFF8800"` (alpha+rgb) | `"ff8800"` (alpha dropped) |
| `"f80"` (CSS short) | `"ff8800"` |
| `""` / nullish | `undefined` |

`renderColor(rrggbb): "#rrggbb"` is the inverse for inline styles.

## Highlight semantics per format

| Format | Capability | Wire mapping |
|---|---|---|
| DOCX | `native` | `highlight: "ffff00"` → nearest match in `HIGHLIGHT_PALETTE` → `w:highlight w:val="yellow"` |
| XLSX | `fill-fallback` | `highlight: "ffff00"` → `xlsx:set-cell-format` with `fill: { color: "ffff00", pattern: "solid" }` |
| PPTX | `native` | `highlight: "ffff00"` → `<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>` |

When `capabilities.highlight === "none"` the picker is hidden. The
shared `HighlightPicker` accepts an arbitrary RRGGBB plus a "clear"
button — DOCX adapter quantises onto the OOXML enum.

## MIXED semantics

`MaybeMixed<T>` collapses an iterable of values:

- All values strictly equal → that value.
- All values are `undefined` → `undefined`.
- Otherwise → `MIXED`.

The picker primitives interpret `MIXED` as "show a `—` placeholder,
preserve the popover state". Applying any picker click sends the
clicked value as a deterministic patch — there is no explicit
"toggle" semantics in the shared layer; the provider observes its
own state on each render and decides.

## What this spec does NOT cover

- Paragraph-level formatting (alignment, indent, line spacing,
  lists). Those stay product-specific until a future "block
  formatting" wave.
- Cell-only properties (number formats, borders, vertical alignment,
  wrap text). Those remain in the XLSX-specific toolbar.
- Hyperlinks, comments, tracked changes, revision marks. They are
  structural, not presentation.
- Theme color references (`w:asciiTheme`, `a:schemeClr`). Adapters
  resolve these to literal `fontFamily` / `color` for the picker but
  do not let the picker mutate them — clearing or re-setting either
  through the shared UI replaces the reference with a literal.

## Reading order

1. `packages/text-formatting/src/types.ts` — the contract.
2. `apps/web/app/editor/docxFormatProvider.ts` — DOCX adapter.
3. `apps/web/app/xlsx-editor/xlsxFormatProvider.ts` — XLSX adapter.
4. `apps/web/app/pptx-editor/pptxFormatProvider.ts` — PPTX adapter.
5. `packages/ui/src/primitives/text-format-bar.tsx` — the composed
   bar that consumes a provider.
