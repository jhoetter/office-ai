import { ooxml } from "@officeai/core";
import {
  defaultStyleTable,
  type StyleAlignment,
  type StyleBorder,
  type StyleBorderSide,
  type StyleCellXf,
  type StyleColor,
  type StyleFill,
  type StyleFont,
  type StyleProtection,
  type StyleTable,
} from "../model/style-table.js";
import { XlsxParseError } from "./errors.js";

const STYLES_PART = "xl/styles.xml";

const FONT_MODELED_TAGS = new Set([
  "name",
  "sz",
  "b",
  "i",
  "u",
  "strike",
  "color",
  "family",
  "scheme",
  "charset",
]);

const BORDER_MODELED_TAGS = new Set(["left", "right", "top", "bottom", "diagonal"]);

/**
 * Parse `xl/styles.xml` text into a typed `StyleTable`.
 * The serializer in `serializer/styles.ts` is the inverse — round-trip
 * is semantic, not byte-identical (attribute order can differ).
 */
export function parseStylesXml(xml: string): StyleTable {
  let tree: unknown;
  try {
    tree = ooxml.parseXml(xml);
  } catch (err) {
    throw new XlsxParseError("invalid-xml", `Failed to parse ${STYLES_PART}`, {
      partPath: STYLES_PART,
      cause: err,
    });
  }

  if (!Array.isArray(tree)) {
    throw new XlsxParseError("invalid-styles", `${STYLES_PART} root must be an element list`, {
      partPath: STYLES_PART,
    });
  }

  const root = (tree as unknown[])
    .map((n) => ooxml.asElement(n))
    .find((el): el is ooxml.XmlElement => el !== null && el.tag === "styleSheet");
  if (!root) {
    throw new XlsxParseError("invalid-styles", `${STYLES_PART} missing <styleSheet> root`, {
      partPath: STYLES_PART,
    });
  }

  const rootAttrs: Record<string, string> = {};
  for (const [k, v] of Object.entries(root.attrs)) {
    if (k === "xmlns") continue;
    rootAttrs[k] = v;
  }

  const numFmts = parseNumFmts(root);
  const fonts = parseFonts(root);
  const fills = parseFills(root);
  const borders = parseBorders(root);
  const cellStyleXfs = parseXfs(root, "cellStyleXfs");
  const cellXfs = parseXfs(root, "cellXfs");

  const cellStylesXml = serializeOpaqueSection(root, "cellStyles");
  const dxfsXml = serializeOpaqueSection(root, "dxfs");
  const tableStylesXml = serializeOpaqueSection(root, "tableStyles");
  const colorsXml = serializeOpaqueSection(root, "colors");
  const extLstXml = serializeOpaqueSection(root, "extLst");

  return {
    numFmts,
    fonts,
    fills,
    borders,
    cellStyleXfs,
    cellXfs,
    cellStylesXml,
    dxfsXml,
    tableStylesXml,
    colorsXml,
    extLstXml,
    rootAttrs,
  };
}

export { defaultStyleTable, STYLES_PART };

function parseNumFmts(root: ooxml.XmlElement): ReadonlyMap<number, string> {
  const out = new Map<number, string>();
  const section = ooxml.findChild(root.children, "numFmts");
  if (!section) return out;
  for (const fmt of ooxml.filterChildren(section.children, "numFmt")) {
    const id = parseIntAttr(fmt.attrs.numFmtId);
    const code = fmt.attrs.formatCode;
    if (id === undefined || code === undefined) continue;
    out.set(id, code);
  }
  return out;
}

function parseFonts(root: ooxml.XmlElement): ReadonlyArray<StyleFont> {
  const section = ooxml.findChild(root.children, "fonts");
  if (!section) return [];
  return ooxml.filterChildren(section.children, "font").map(parseFont);
}

function parseFont(el: ooxml.XmlElement): StyleFont {
  const font: {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: string | true;
    strike?: boolean;
    color?: StyleColor;
    family?: number;
    scheme?: string;
    charset?: number;
    opaqueExtras: string[];
  } = { opaqueExtras: [] };

  for (const child of el.children) {
    const c = ooxml.asElement(child);
    if (!c) continue;
    if (!FONT_MODELED_TAGS.has(c.tag)) {
      font.opaqueExtras.push(serializeNode(child));
      continue;
    }
    switch (c.tag) {
      case "name":
        if (c.attrs.val !== undefined) font.name = c.attrs.val;
        break;
      case "sz": {
        const n = parseFloatAttr(c.attrs.val);
        if (n !== undefined) font.size = n;
        break;
      }
      case "b":
        font.bold = parseBoolAttr(c.attrs.val, true);
        break;
      case "i":
        font.italic = parseBoolAttr(c.attrs.val, true);
        break;
      case "u":
        font.underline = c.attrs.val ?? true;
        break;
      case "strike":
        font.strike = parseBoolAttr(c.attrs.val, true);
        break;
      case "color":
        font.color = parseColor(c);
        break;
      case "family": {
        const n = parseIntAttr(c.attrs.val);
        if (n !== undefined) font.family = n;
        break;
      }
      case "scheme":
        if (c.attrs.val !== undefined) font.scheme = c.attrs.val;
        break;
      case "charset": {
        const n = parseIntAttr(c.attrs.val);
        if (n !== undefined) font.charset = n;
        break;
      }
    }
  }

  return font;
}

function parseFills(root: ooxml.XmlElement): ReadonlyArray<StyleFill> {
  const section = ooxml.findChild(root.children, "fills");
  if (!section) return [];
  return ooxml.filterChildren(section.children, "fill").map(parseFill);
}

function parseFill(el: ooxml.XmlElement): StyleFill {
  for (const child of el.children) {
    const c = ooxml.asElement(child);
    if (!c) continue;
    if (c.tag === "patternFill") {
      const fill: {
        kind: "pattern";
        patternType?: string;
        fgColor?: StyleColor;
        bgColor?: StyleColor;
      } = { kind: "pattern" };
      if (c.attrs.patternType !== undefined) fill.patternType = c.attrs.patternType;
      const fg = ooxml.findChild(c.children, "fgColor");
      const bg = ooxml.findChild(c.children, "bgColor");
      if (fg) fill.fgColor = parseColor(fg);
      if (bg) fill.bgColor = parseColor(bg);
      return fill;
    }
    if (c.tag === "gradientFill") {
      return { kind: "gradient", gradientXml: serializeNode(child) };
    }
  }
  return { kind: "pattern", patternType: "none" };
}

function parseBorders(root: ooxml.XmlElement): ReadonlyArray<StyleBorder> {
  const section = ooxml.findChild(root.children, "borders");
  if (!section) return [];
  return ooxml.filterChildren(section.children, "border").map(parseBorder);
}

function parseBorder(el: ooxml.XmlElement): StyleBorder {
  const border: {
    left?: StyleBorderSide;
    right?: StyleBorderSide;
    top?: StyleBorderSide;
    bottom?: StyleBorderSide;
    diagonal?: StyleBorderSide;
    diagonalUp?: boolean;
    diagonalDown?: boolean;
    outline?: boolean;
    opaqueExtras: string[];
  } = { opaqueExtras: [] };
  if (el.attrs.diagonalUp !== undefined) border.diagonalUp = parseBoolAttr(el.attrs.diagonalUp, false);
  if (el.attrs.diagonalDown !== undefined) border.diagonalDown = parseBoolAttr(el.attrs.diagonalDown, false);
  if (el.attrs.outline !== undefined) border.outline = parseBoolAttr(el.attrs.outline, false);

  for (const child of el.children) {
    const c = ooxml.asElement(child);
    if (!c) continue;
    if (!BORDER_MODELED_TAGS.has(c.tag)) {
      border.opaqueExtras.push(serializeNode(child));
      continue;
    }
    const side = parseBorderSide(c);
    switch (c.tag) {
      case "left":
        border.left = side;
        break;
      case "right":
        border.right = side;
        break;
      case "top":
        border.top = side;
        break;
      case "bottom":
        border.bottom = side;
        break;
      case "diagonal":
        border.diagonal = side;
        break;
    }
  }

  return border;
}

function parseBorderSide(el: ooxml.XmlElement): StyleBorderSide {
  const side: { style?: string; color?: StyleColor } = {};
  if (el.attrs.style !== undefined) side.style = el.attrs.style;
  const color = ooxml.findChild(el.children, "color");
  if (color) side.color = parseColor(color);
  return side;
}

function parseXfs(root: ooxml.XmlElement, sectionTag: string): ReadonlyArray<StyleCellXf> {
  const section = ooxml.findChild(root.children, sectionTag);
  if (!section) return [];
  return ooxml.filterChildren(section.children, "xf").map(parseXf);
}

function parseXf(el: ooxml.XmlElement): StyleCellXf {
  const xf: {
    numFmtId?: number;
    fontId?: number;
    fillId?: number;
    borderId?: number;
    xfId?: number;
    applyNumberFormat?: boolean;
    applyFont?: boolean;
    applyFill?: boolean;
    applyBorder?: boolean;
    applyAlignment?: boolean;
    applyProtection?: boolean;
    quotePrefix?: boolean;
    pivotButton?: boolean;
    alignment?: StyleAlignment;
    protection?: StyleProtection;
  } = {};

  if (el.attrs.numFmtId !== undefined) xf.numFmtId = parseIntAttr(el.attrs.numFmtId);
  if (el.attrs.fontId !== undefined) xf.fontId = parseIntAttr(el.attrs.fontId);
  if (el.attrs.fillId !== undefined) xf.fillId = parseIntAttr(el.attrs.fillId);
  if (el.attrs.borderId !== undefined) xf.borderId = parseIntAttr(el.attrs.borderId);
  if (el.attrs.xfId !== undefined) xf.xfId = parseIntAttr(el.attrs.xfId);
  if (el.attrs.applyNumberFormat !== undefined)
    xf.applyNumberFormat = parseBoolAttr(el.attrs.applyNumberFormat, false);
  if (el.attrs.applyFont !== undefined) xf.applyFont = parseBoolAttr(el.attrs.applyFont, false);
  if (el.attrs.applyFill !== undefined) xf.applyFill = parseBoolAttr(el.attrs.applyFill, false);
  if (el.attrs.applyBorder !== undefined) xf.applyBorder = parseBoolAttr(el.attrs.applyBorder, false);
  if (el.attrs.applyAlignment !== undefined)
    xf.applyAlignment = parseBoolAttr(el.attrs.applyAlignment, false);
  if (el.attrs.applyProtection !== undefined)
    xf.applyProtection = parseBoolAttr(el.attrs.applyProtection, false);
  if (el.attrs.quotePrefix !== undefined) xf.quotePrefix = parseBoolAttr(el.attrs.quotePrefix, false);
  if (el.attrs.pivotButton !== undefined) xf.pivotButton = parseBoolAttr(el.attrs.pivotButton, false);

  const alignment = ooxml.findChild(el.children, "alignment");
  if (alignment) xf.alignment = parseAlignment(alignment);
  const protection = ooxml.findChild(el.children, "protection");
  if (protection) xf.protection = parseProtection(protection);

  return xf;
}

function parseAlignment(el: ooxml.XmlElement): StyleAlignment {
  const a: {
    horizontal?: string;
    vertical?: string;
    wrapText?: boolean;
    textRotation?: number;
    indent?: number;
    shrinkToFit?: boolean;
    readingOrder?: number;
    justifyLastLine?: boolean;
    relativeIndent?: number;
  } = {};
  if (el.attrs.horizontal !== undefined) a.horizontal = el.attrs.horizontal;
  if (el.attrs.vertical !== undefined) a.vertical = el.attrs.vertical;
  if (el.attrs.wrapText !== undefined) a.wrapText = parseBoolAttr(el.attrs.wrapText, false);
  if (el.attrs.textRotation !== undefined) a.textRotation = parseIntAttr(el.attrs.textRotation);
  if (el.attrs.indent !== undefined) a.indent = parseIntAttr(el.attrs.indent);
  if (el.attrs.shrinkToFit !== undefined) a.shrinkToFit = parseBoolAttr(el.attrs.shrinkToFit, false);
  if (el.attrs.readingOrder !== undefined) a.readingOrder = parseIntAttr(el.attrs.readingOrder);
  if (el.attrs.justifyLastLine !== undefined)
    a.justifyLastLine = parseBoolAttr(el.attrs.justifyLastLine, false);
  if (el.attrs.relativeIndent !== undefined) a.relativeIndent = parseIntAttr(el.attrs.relativeIndent);
  return a;
}

function parseProtection(el: ooxml.XmlElement): StyleProtection {
  const p: { locked?: boolean; hidden?: boolean } = {};
  if (el.attrs.locked !== undefined) p.locked = parseBoolAttr(el.attrs.locked, true);
  if (el.attrs.hidden !== undefined) p.hidden = parseBoolAttr(el.attrs.hidden, false);
  return p;
}

function parseColor(el: ooxml.XmlElement): StyleColor {
  const c: { rgb?: string; theme?: number; tint?: number; indexed?: number; auto?: boolean } = {};
  if (el.attrs.rgb !== undefined) c.rgb = el.attrs.rgb;
  if (el.attrs.theme !== undefined) c.theme = parseIntAttr(el.attrs.theme);
  if (el.attrs.tint !== undefined) c.tint = parseFloatAttr(el.attrs.tint);
  if (el.attrs.indexed !== undefined) c.indexed = parseIntAttr(el.attrs.indexed);
  if (el.attrs.auto !== undefined) c.auto = parseBoolAttr(el.attrs.auto, false);
  return c;
}

function serializeOpaqueSection(root: ooxml.XmlElement, tag: string): string {
  const section = ooxml.findChild(root.children, tag);
  if (!section) return "";
  return serializeNode(section.entry);
}

function serializeNode(node: unknown): string {
  return ooxml.serializeXml([node], { xmlDeclaration: null });
}

function parseIntAttr(v: string | undefined): number | undefined {
  if (v === undefined) return undefined;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n : undefined;
}

function parseFloatAttr(v: string | undefined): number | undefined {
  if (v === undefined) return undefined;
  const n = Number.parseFloat(v);
  return Number.isFinite(n) ? n : undefined;
}

function parseBoolAttr(v: string | undefined, defaultIfBare: boolean): boolean {
  if (v === undefined) return defaultIfBare;
  return v === "1" || v === "true" || v === "TRUE";
}
