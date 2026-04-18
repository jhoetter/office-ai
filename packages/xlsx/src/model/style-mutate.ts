import type {
  StyleAlignment,
  StyleBorder,
  StyleBorderSide,
  StyleCellXf,
  StyleColor,
  StyleFill,
  StyleFont,
  StyleTable,
} from "./style-table.js";

/**
 * Style intern + content-hash dedupe.
 *
 * The `xlsx:set-cell-format` handler mutates per-cell styles by:
 *
 *   1. Resolving the cell's current `xfId` to a fully-flattened
 *      `EffectiveStyle` (so the per-leaf `applyX` flags don't matter
 *      to the agent — they're auto-set to `true` for any aspect we
 *      override).
 *   2. Merging an agent-facing `CellFormat` patch onto that effective
 *      style.
 *   3. Re-interning the result through this module's content-hash
 *      maps. Identical fonts / fills / borders / xfs are reused so a
 *      bulk format-range only allocates O(distinct prior styles) new
 *      `xfId`s — see §4 spec "small constant" property.
 *
 * The `StyleTable` is conceptually immutable: every helper returns a
 * new table with appended arrays. Existing indices remain valid so we
 * never have to rewrite `Cell.styleId` on a different cell as a side
 * effect of one cell's mutation.
 */

/** Fully-resolved per-cell effective style; the agent's mental model. */
export interface EffectiveStyle {
  readonly numFmtId: number;
  readonly font: StyleFont;
  readonly fill: StyleFill;
  readonly border: StyleBorder;
  readonly alignment?: StyleAlignment;
  readonly protection?: { readonly locked?: boolean; readonly hidden?: boolean };
}

const DEFAULT_FONT: StyleFont = { name: "Calibri", size: 11, opaqueExtras: [] };
const DEFAULT_FILL: StyleFill = { kind: "pattern", patternType: "none" };
const DEFAULT_BORDER: StyleBorder = {
  left: {},
  right: {},
  top: {},
  bottom: {},
  diagonal: {},
  opaqueExtras: [],
};

/** Flatten a `cellXfs` row into its concrete font/fill/border. */
export function flattenCellXf(table: StyleTable, xfId: number | undefined): EffectiveStyle {
  const xf = xfId !== undefined ? table.cellXfs[xfId] : undefined;
  return resolveXf(table, xf);
}

function resolveXf(table: StyleTable, xf: StyleCellXf | undefined): EffectiveStyle {
  const numFmtId = xf?.numFmtId ?? 0;
  const font = xf?.fontId !== undefined ? (table.fonts[xf.fontId] ?? DEFAULT_FONT) : DEFAULT_FONT;
  const fill = xf?.fillId !== undefined ? (table.fills[xf.fillId] ?? DEFAULT_FILL) : DEFAULT_FILL;
  const border = xf?.borderId !== undefined ? (table.borders[xf.borderId] ?? DEFAULT_BORDER) : DEFAULT_BORDER;
  const eff: {
    numFmtId: number;
    font: StyleFont;
    fill: StyleFill;
    border: StyleBorder;
    alignment?: StyleAlignment;
    protection?: { locked?: boolean; hidden?: boolean };
  } = { numFmtId, font, fill, border };
  if (xf?.alignment) eff.alignment = xf.alignment;
  if (xf?.protection) eff.protection = xf.protection;
  return eff;
}

/** Result of `internStyle`. The new table + the resolved per-cell xfId. */
export interface InternResult {
  readonly table: StyleTable;
  readonly xfId: number;
  /** Indices that were freshly appended (for `style-added` diff entries). */
  readonly addedFontId?: number;
  readonly addedFillId?: number;
  readonly addedBorderId?: number;
  readonly addedNumFmtId?: number;
  readonly addedXfId?: number;
}

/**
 * Intern an effective style into the table. Returns the same table
 * (with the same xfId) when every component already exists; otherwise
 * appends new component entries and a new xf row. The `applyX` flags
 * on the new xf are set to `true` for every aspect we model so Excel
 * picks up the cell's own values rather than inheriting from the
 * named-style parent.
 */
export function internStyle(table: StyleTable, eff: EffectiveStyle): InternResult {
  const fontIntern = internFont(table, eff.font);
  const fillIntern = internFill(fontIntern.table, eff.fill);
  const borderIntern = internBorder(fillIntern.table, eff.border);

  const xf: StyleCellXf = {
    numFmtId: eff.numFmtId,
    fontId: fontIntern.id,
    fillId: fillIntern.id,
    borderId: borderIntern.id,
    xfId: 0,
    applyNumberFormat: true,
    applyFont: true,
    applyFill: true,
    applyBorder: true,
    ...(eff.alignment ? { applyAlignment: true, alignment: eff.alignment } : {}),
    ...(eff.protection ? { applyProtection: true, protection: eff.protection } : {}),
  };

  const xfHash = hashXf(xf);
  for (let i = 0; i < borderIntern.table.cellXfs.length; i++) {
    if (hashXf(borderIntern.table.cellXfs[i]) === xfHash) {
      const result: {
        table: StyleTable;
        xfId: number;
        addedFontId?: number;
        addedFillId?: number;
        addedBorderId?: number;
      } = { table: borderIntern.table, xfId: i };
      if (fontIntern.added) result.addedFontId = fontIntern.id;
      if (fillIntern.added) result.addedFillId = fillIntern.id;
      if (borderIntern.added) result.addedBorderId = borderIntern.id;
      return result;
    }
  }

  const cellXfs = [...borderIntern.table.cellXfs, xf];
  const nextTable: StyleTable = { ...borderIntern.table, cellXfs };
  const result: {
    table: StyleTable;
    xfId: number;
    addedXfId: number;
    addedFontId?: number;
    addedFillId?: number;
    addedBorderId?: number;
  } = {
    table: nextTable,
    xfId: cellXfs.length - 1,
    addedXfId: cellXfs.length - 1,
  };
  if (fontIntern.added) result.addedFontId = fontIntern.id;
  if (fillIntern.added) result.addedFillId = fillIntern.id;
  if (borderIntern.added) result.addedBorderId = borderIntern.id;
  return result;
}

/** Append a custom number format starting at `nextCustomId`. Returns id + table. */
export function internNumberFormat(
  table: StyleTable,
  formatCode: string
): { table: StyleTable; numFmtId: number; added: boolean } {
  for (const [id, code] of table.numFmts) {
    if (code === formatCode) return { table, numFmtId: id, added: false };
  }
  let nextId = 164;
  for (const id of table.numFmts.keys()) {
    if (id >= nextId) nextId = id + 1;
  }
  const numFmts = new Map(table.numFmts);
  numFmts.set(nextId, formatCode);
  return { table: { ...table, numFmts }, numFmtId: nextId, added: true };
}

interface InternComponent {
  readonly table: StyleTable;
  readonly id: number;
  readonly added: boolean;
}

function internFont(table: StyleTable, font: StyleFont): InternComponent {
  const hash = hashFont(font);
  for (let i = 0; i < table.fonts.length; i++) {
    if (hashFont(table.fonts[i]) === hash) return { table, id: i, added: false };
  }
  const fonts = [...table.fonts, font];
  return { table: { ...table, fonts }, id: fonts.length - 1, added: true };
}

function internFill(table: StyleTable, fill: StyleFill): InternComponent {
  const hash = hashFill(fill);
  for (let i = 0; i < table.fills.length; i++) {
    if (hashFill(table.fills[i]) === hash) return { table, id: i, added: false };
  }
  const fills = [...table.fills, fill];
  return { table: { ...table, fills }, id: fills.length - 1, added: true };
}

function internBorder(table: StyleTable, border: StyleBorder): InternComponent {
  const hash = hashBorder(border);
  for (let i = 0; i < table.borders.length; i++) {
    if (hashBorder(table.borders[i]) === hash) return { table, id: i, added: false };
  }
  const borders = [...table.borders, border];
  return { table: { ...table, borders }, id: borders.length - 1, added: true };
}

function hashFont(font: StyleFont): string {
  return JSON.stringify({
    name: font.name,
    size: font.size,
    bold: font.bold,
    italic: font.italic,
    underline: font.underline,
    strike: font.strike,
    color: canonicalColor(font.color),
    family: font.family,
    scheme: font.scheme,
    charset: font.charset,
    extras: font.opaqueExtras,
  });
}

function hashFill(fill: StyleFill): string {
  return JSON.stringify({
    kind: fill.kind,
    patternType: fill.patternType,
    fgColor: canonicalColor(fill.fgColor),
    bgColor: canonicalColor(fill.bgColor),
    gradientXml: fill.gradientXml,
  });
}

function hashBorder(border: StyleBorder): string {
  return JSON.stringify({
    left: canonicalSide(border.left),
    right: canonicalSide(border.right),
    top: canonicalSide(border.top),
    bottom: canonicalSide(border.bottom),
    diagonal: canonicalSide(border.diagonal),
    diagonalUp: border.diagonalUp,
    diagonalDown: border.diagonalDown,
    outline: border.outline,
    extras: border.opaqueExtras,
  });
}

function hashXf(xf: StyleCellXf): string {
  return JSON.stringify({
    numFmtId: xf.numFmtId,
    fontId: xf.fontId,
    fillId: xf.fillId,
    borderId: xf.borderId,
    xfId: xf.xfId,
    applyNumberFormat: xf.applyNumberFormat,
    applyFont: xf.applyFont,
    applyFill: xf.applyFill,
    applyBorder: xf.applyBorder,
    applyAlignment: xf.applyAlignment,
    applyProtection: xf.applyProtection,
    quotePrefix: xf.quotePrefix,
    pivotButton: xf.pivotButton,
    alignment: canonicalAlignment(xf.alignment),
    protection: xf.protection,
  });
}

function canonicalSide(side: StyleBorderSide | undefined): unknown {
  if (!side) return undefined;
  return { style: side.style, color: canonicalColor(side.color) };
}

function canonicalColor(color: StyleColor | undefined): unknown {
  if (!color) return undefined;
  return {
    rgb: color.rgb,
    theme: color.theme,
    tint: color.tint,
    indexed: color.indexed,
    auto: color.auto,
  };
}

function canonicalAlignment(a: StyleAlignment | undefined): unknown {
  if (!a) return undefined;
  return {
    horizontal: a.horizontal,
    vertical: a.vertical,
    wrapText: a.wrapText,
    textRotation: a.textRotation,
    indent: a.indent,
    shrinkToFit: a.shrinkToFit,
    readingOrder: a.readingOrder,
    justifyLastLine: a.justifyLastLine,
    relativeIndent: a.relativeIndent,
  };
}
