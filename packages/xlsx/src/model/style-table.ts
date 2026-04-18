/**
 * Typed model of `xl/styles.xml`.
 *
 * The shape mirrors `spec/xlsx/document-model.md` §6 but trimmed to
 * what the agent actually authors in P0:
 *
 *   - **Modeled in detail:** `numFmts`, `fonts`, `fills`, `borders`,
 *     `cellStyleXfs`, `cellXfs`. These are the rows the
 *     `xlsx:set-cell-format` handler reads, hashes, and appends to.
 *   - **Preserved as opaque XML:** `cellStyles`, `dxfs`, `tableStyles`,
 *     `colors`, `extLst`. We round-trip these byte-for-byte by holding
 *     the original element source. The agent does not author them.
 *
 * Sub-elements we don't model on a font/fill/border (e.g. `<scheme>`,
 * `<vertAlign>`, `<extLst>`) are kept on a per-record `opaqueExtras`
 * string array so the re-emit slots them back in.
 *
 * The table is conceptually immutable; mutation helpers in
 * `model/style-mutate.ts` return a new `StyleTable`.
 */

export interface StyleAlignment {
  readonly horizontal?: string;
  readonly vertical?: string;
  readonly wrapText?: boolean;
  readonly textRotation?: number;
  readonly indent?: number;
  readonly shrinkToFit?: boolean;
  readonly readingOrder?: number;
  readonly justifyLastLine?: boolean;
  readonly relativeIndent?: number;
}

export interface StyleProtection {
  readonly locked?: boolean;
  readonly hidden?: boolean;
}

/** Raw OOXML colour element. We pass `<color>`'s attributes through. */
export interface StyleColor {
  readonly rgb?: string;
  readonly theme?: number;
  readonly tint?: number;
  readonly indexed?: number;
  readonly auto?: boolean;
}

export interface StyleFont {
  readonly name?: string;
  readonly size?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: string | true;
  readonly strike?: boolean;
  readonly color?: StyleColor;
  readonly family?: number;
  readonly scheme?: string;
  readonly charset?: number;
  /** Verbatim XML for `<vertAlign>`, `<extLst>`, … we don't otherwise model. */
  readonly opaqueExtras: ReadonlyArray<string>;
}

export interface StyleFill {
  /** `"pattern"` (incl. solid/none) or `"gradient"`. */
  readonly kind: "pattern" | "gradient";
  readonly patternType?: string;
  readonly fgColor?: StyleColor;
  readonly bgColor?: StyleColor;
  /** For gradient fills we keep the `<gradientFill>` element verbatim. */
  readonly gradientXml?: string;
}

export interface StyleBorderSide {
  readonly style?: string;
  readonly color?: StyleColor;
}

export interface StyleBorder {
  readonly left?: StyleBorderSide;
  readonly right?: StyleBorderSide;
  readonly top?: StyleBorderSide;
  readonly bottom?: StyleBorderSide;
  readonly diagonal?: StyleBorderSide;
  readonly diagonalUp?: boolean;
  readonly diagonalDown?: boolean;
  readonly outline?: boolean;
  /** Verbatim XML for `<vertical>`, `<horizontal>`, `<extLst>`. */
  readonly opaqueExtras: ReadonlyArray<string>;
}

export interface StyleCellXf {
  readonly numFmtId?: number;
  readonly fontId?: number;
  readonly fillId?: number;
  readonly borderId?: number;
  readonly xfId?: number;
  readonly applyNumberFormat?: boolean;
  readonly applyFont?: boolean;
  readonly applyFill?: boolean;
  readonly applyBorder?: boolean;
  readonly applyAlignment?: boolean;
  readonly applyProtection?: boolean;
  readonly quotePrefix?: boolean;
  readonly pivotButton?: boolean;
  readonly alignment?: StyleAlignment;
  readonly protection?: StyleProtection;
}

export interface StyleNumberFormat {
  readonly numFmtId: number;
  readonly formatCode: string;
}

export interface StyleTable {
  /**
   * Custom number formats keyed by `numFmtId`. IDs 0–49 are Excel
   * built-ins and never appear here; IDs ≥ 164 are user-registered.
   * Some legacy workbooks register IDs in the 50–163 range; we keep
   * those verbatim.
   */
  readonly numFmts: ReadonlyMap<number, string>;
  readonly fonts: ReadonlyArray<StyleFont>;
  readonly fills: ReadonlyArray<StyleFill>;
  readonly borders: ReadonlyArray<StyleBorder>;
  /** Named-style xfs (`<cellStyleXfs>`); referenced by `xfId` on `cellXfs`. */
  readonly cellStyleXfs: ReadonlyArray<StyleCellXf>;
  /** Per-cell xfs (`<cellXfs>`); `Cell.styleId` is an index into this. */
  readonly cellXfs: ReadonlyArray<StyleCellXf>;
  /**
   * Top-level sections we preserve verbatim. Empty string = the
   * section was not present in the source. Re-emitted as-is when set.
   */
  readonly cellStylesXml: string;
  readonly dxfsXml: string;
  readonly tableStylesXml: string;
  readonly colorsXml: string;
  readonly extLstXml: string;
  /**
   * Root `<styleSheet>` attributes (xmlns + namespace decls). Re-emitted
   * verbatim. The default `xmlns` is always written by the serializer,
   * so it may be absent here.
   */
  readonly rootAttrs: Readonly<Record<string, string>>;
}

/**
 * Default StyleTable for workbooks lacking `xl/styles.xml`. Mirrors
 * what Excel implicitly assumes: one default font, two stock fills
 * (`none` + `gray125`), one empty border, and one no-op cellXf.
 */
export function defaultStyleTable(): StyleTable {
  const defaultFont: StyleFont = {
    name: "Calibri",
    size: 11,
    color: { theme: 1 },
    family: 2,
    scheme: "minor",
    opaqueExtras: [],
  };
  const noneFill: StyleFill = { kind: "pattern", patternType: "none" };
  const gray125Fill: StyleFill = { kind: "pattern", patternType: "gray125" };
  const emptyBorder: StyleBorder = {
    left: {},
    right: {},
    top: {},
    bottom: {},
    diagonal: {},
    opaqueExtras: [],
  };
  const defaultXf: StyleCellXf = {
    numFmtId: 0,
    fontId: 0,
    fillId: 0,
    borderId: 0,
    xfId: 0,
  };
  return {
    numFmts: new Map(),
    fonts: [defaultFont],
    fills: [noneFill, gray125Fill],
    borders: [emptyBorder],
    cellStyleXfs: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }],
    cellXfs: [defaultXf],
    cellStylesXml: "",
    dxfsXml: "",
    tableStylesXml: "",
    colorsXml: "",
    extLstXml: "",
    rootAttrs: {},
  };
}
