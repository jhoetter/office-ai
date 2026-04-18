/**
 * Canonical, format-agnostic vocabulary for run-level text formatting.
 *
 * All units normalised so that consumers (UI pickers, providers) speak
 * one language. Format adapters convert into native units when
 * dispatching commands to their respective bus.
 *
 * - Font sizes: points (float)
 * - Colors: 6-char lowercase RRGGBB hex, no leading '#'
 * - Underline: boolean (single) or one of UnderlineStyle for variants
 *
 * See spec/shared/text-formatting.md for the wider contract.
 */

export type UnderlineStyle =
  | "single"
  | "double"
  | "thick"
  | "dotted"
  | "dashed"
  | "wavy";

/**
 * A patch of formatting to apply to the current selection. Every
 * field is optional; only the supplied fields are applied. An empty
 * string for `color` / `highlight` / `fontFamily` is interpreted by
 * adapters as "clear this attribute".
 */
export interface TextFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | UnderlineStyle;
  strike?: boolean;
  fontFamily?: string;
  /** Canonical font size in points. */
  fontSizePt?: number;
  /** Lowercase RRGGBB, no '#'. Empty string means "clear". */
  color?: string;
  /**
   * Lowercase RRGGBB, no '#'. Empty string means "clear".
   * Adapters map this onto the most appropriate native concept
   * (DOCX w:highlight enum / XLSX cell fill / PPTX a:highlight).
   */
  highlight?: string;
}

/**
 * Sentinel for "selection has more than one value for this attribute".
 * The pickers render `MIXED` as a `—` placeholder.
 */
export const MIXED: unique symbol = Symbol("text-formatting/MIXED");
export type Mixed = typeof MIXED;

/**
 * Possible value of an attribute when read across the selection:
 * - `T`: every cell/run agrees on this value
 * - `undefined`: every cell/run is unset (inheriting default)
 * - `MIXED`: at least two values disagree
 */
export type MaybeMixed<T> = T | Mixed | undefined;

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

export interface TextFormatCapabilities {
  /**
   * - `native`: round-trips as a real character highlight
   *   (DOCX w:highlight, PPTX a:highlight)
   * - `fill-fallback`: shared HighlightPicker dispatches a cell-fill
   *   patch instead (XLSX)
   * - `none`: hide the highlight picker
   */
  readonly highlight: "native" | "fill-fallback" | "none";
  readonly underlineVariants: boolean;
  readonly fontFamily: boolean;
  readonly fontSize: boolean;
  readonly strike: boolean;
}

/**
 * The contract every editor implements to wire its native model into
 * the shared TextFormatBar. Adapters live in the web app (e.g.
 * `apps/web/app/editor/docxFormatProvider.ts`).
 */
export interface TextFormatProvider {
  /** Cheap; called on every render. */
  getActive(): ActiveTextFormat;
  /** Dispatches the format-specific command via the agent. */
  apply(patch: TextFormat): void;
  /** True iff there is anything to format. */
  hasSelection(): boolean;
  readonly capabilities: TextFormatCapabilities;
}
