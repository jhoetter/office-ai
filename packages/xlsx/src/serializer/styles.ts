import type {
  StyleAlignment,
  StyleBorder,
  StyleBorderSide,
  StyleCellXf,
  StyleColor,
  StyleFill,
  StyleFont,
  StyleProtection,
  StyleTable,
} from "../model/style-table.js";

const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/**
 * Re-emit `xl/styles.xml` from a typed `StyleTable`.
 *
 * Round-trip is **semantic**, not byte-identical: attribute order is
 * not preserved through `parseXml → builder.build`, so we emit a
 * canonical attribute order. The parser in `parser/styles.ts` is the
 * inverse: re-parsing this output yields a structurally equivalent
 * `StyleTable`.
 *
 * Top-level sections (`cellStyles`, `dxfs`, `tableStyles`, `colors`,
 * `extLst`) round-trip verbatim because we hold their original element
 * source on `StyleTable.{section}Xml` rather than modeling them.
 */
export function serializeStylesXml(table: StyleTable): string {
  const parts: string[] = [];
  parts.push(XML_DECL);
  parts.push("<styleSheet");
  parts.push(` xmlns="${NS}"`);
  for (const [k, v] of Object.entries(table.rootAttrs)) {
    parts.push(` ${k}="${escapeAttr(v)}"`);
  }
  parts.push(">");

  if (table.numFmts.size > 0) {
    parts.push(`<numFmts count="${table.numFmts.size}">`);
    const ids = Array.from(table.numFmts.keys()).sort((a, b) => a - b);
    for (const id of ids) {
      const code = table.numFmts.get(id) ?? "";
      parts.push(`<numFmt numFmtId="${id}" formatCode="${escapeAttr(code)}"/>`);
    }
    parts.push("</numFmts>");
  }

  parts.push(`<fonts count="${table.fonts.length}">`);
  for (const f of table.fonts) parts.push(emitFont(f));
  parts.push("</fonts>");

  parts.push(`<fills count="${table.fills.length}">`);
  for (const fill of table.fills) parts.push(emitFill(fill));
  parts.push("</fills>");

  parts.push(`<borders count="${table.borders.length}">`);
  for (const b of table.borders) parts.push(emitBorder(b));
  parts.push("</borders>");

  parts.push(`<cellStyleXfs count="${table.cellStyleXfs.length}">`);
  for (const xf of table.cellStyleXfs) parts.push(emitXf(xf));
  parts.push("</cellStyleXfs>");

  parts.push(`<cellXfs count="${table.cellXfs.length}">`);
  for (const xf of table.cellXfs) parts.push(emitXf(xf));
  parts.push("</cellXfs>");

  if (table.cellStylesXml) parts.push(table.cellStylesXml);
  if (table.dxfsXml) parts.push(table.dxfsXml);
  if (table.tableStylesXml) parts.push(table.tableStylesXml);
  if (table.colorsXml) parts.push(table.colorsXml);
  if (table.extLstXml) parts.push(table.extLstXml);

  parts.push("</styleSheet>");
  return parts.join("");
}

function emitFont(f: StyleFont): string {
  const parts: string[] = ["<font>"];
  if (f.bold !== undefined) parts.push(emitBoolEl("b", f.bold));
  if (f.italic !== undefined) parts.push(emitBoolEl("i", f.italic));
  if (f.strike !== undefined) parts.push(emitBoolEl("strike", f.strike));
  if (f.underline !== undefined) {
    if (f.underline === true) parts.push("<u/>");
    else parts.push(`<u val="${escapeAttr(String(f.underline))}"/>`);
  }
  if (f.size !== undefined) parts.push(`<sz val="${formatNumber(f.size)}"/>`);
  if (f.color) parts.push(emitColor("color", f.color));
  if (f.name !== undefined) parts.push(`<name val="${escapeAttr(f.name)}"/>`);
  if (f.family !== undefined) parts.push(`<family val="${f.family}"/>`);
  if (f.charset !== undefined) parts.push(`<charset val="${f.charset}"/>`);
  if (f.scheme !== undefined) parts.push(`<scheme val="${escapeAttr(f.scheme)}"/>`);
  for (const extra of f.opaqueExtras) parts.push(extra);
  parts.push("</font>");
  return parts.join("");
}

function emitFill(fill: StyleFill): string {
  if (fill.kind === "gradient") {
    return `<fill>${fill.gradientXml ?? "<gradientFill/>"}</fill>`;
  }
  const attrs = fill.patternType !== undefined ? ` patternType="${escapeAttr(fill.patternType)}"` : "";
  const inner: string[] = [];
  if (fill.fgColor) inner.push(emitColor("fgColor", fill.fgColor));
  if (fill.bgColor) inner.push(emitColor("bgColor", fill.bgColor));
  if (inner.length === 0) {
    return `<fill><patternFill${attrs}/></fill>`;
  }
  return `<fill><patternFill${attrs}>${inner.join("")}</patternFill></fill>`;
}

function emitBorder(b: StyleBorder): string {
  const attrs: string[] = [];
  if (b.diagonalUp !== undefined) attrs.push(` diagonalUp="${b.diagonalUp ? "1" : "0"}"`);
  if (b.diagonalDown !== undefined) attrs.push(` diagonalDown="${b.diagonalDown ? "1" : "0"}"`);
  if (b.outline !== undefined) attrs.push(` outline="${b.outline ? "1" : "0"}"`);
  const sides: string[] = [];
  sides.push(emitBorderSide("left", b.left));
  sides.push(emitBorderSide("right", b.right));
  sides.push(emitBorderSide("top", b.top));
  sides.push(emitBorderSide("bottom", b.bottom));
  sides.push(emitBorderSide("diagonal", b.diagonal));
  for (const extra of b.opaqueExtras) sides.push(extra);
  return `<border${attrs.join("")}>${sides.join("")}</border>`;
}

function emitBorderSide(tag: string, side: StyleBorderSide | undefined): string {
  if (!side) return `<${tag}/>`;
  const attrs = side.style !== undefined ? ` style="${escapeAttr(side.style)}"` : "";
  if (!side.color) return `<${tag}${attrs}/>`;
  return `<${tag}${attrs}>${emitColor("color", side.color)}</${tag}>`;
}

function emitXf(xf: StyleCellXf): string {
  const attrs: string[] = [];
  if (xf.numFmtId !== undefined) attrs.push(` numFmtId="${xf.numFmtId}"`);
  if (xf.fontId !== undefined) attrs.push(` fontId="${xf.fontId}"`);
  if (xf.fillId !== undefined) attrs.push(` fillId="${xf.fillId}"`);
  if (xf.borderId !== undefined) attrs.push(` borderId="${xf.borderId}"`);
  if (xf.xfId !== undefined) attrs.push(` xfId="${xf.xfId}"`);
  if (xf.applyNumberFormat !== undefined)
    attrs.push(` applyNumberFormat="${xf.applyNumberFormat ? "1" : "0"}"`);
  if (xf.applyFont !== undefined) attrs.push(` applyFont="${xf.applyFont ? "1" : "0"}"`);
  if (xf.applyFill !== undefined) attrs.push(` applyFill="${xf.applyFill ? "1" : "0"}"`);
  if (xf.applyBorder !== undefined) attrs.push(` applyBorder="${xf.applyBorder ? "1" : "0"}"`);
  if (xf.applyAlignment !== undefined) attrs.push(` applyAlignment="${xf.applyAlignment ? "1" : "0"}"`);
  if (xf.applyProtection !== undefined) attrs.push(` applyProtection="${xf.applyProtection ? "1" : "0"}"`);
  if (xf.quotePrefix !== undefined) attrs.push(` quotePrefix="${xf.quotePrefix ? "1" : "0"}"`);
  if (xf.pivotButton !== undefined) attrs.push(` pivotButton="${xf.pivotButton ? "1" : "0"}"`);

  const inner: string[] = [];
  if (xf.alignment) inner.push(emitAlignment(xf.alignment));
  if (xf.protection) inner.push(emitProtection(xf.protection));

  if (inner.length === 0) return `<xf${attrs.join("")}/>`;
  return `<xf${attrs.join("")}>${inner.join("")}</xf>`;
}

function emitAlignment(a: StyleAlignment): string {
  const attrs: string[] = [];
  if (a.horizontal !== undefined) attrs.push(` horizontal="${escapeAttr(a.horizontal)}"`);
  if (a.vertical !== undefined) attrs.push(` vertical="${escapeAttr(a.vertical)}"`);
  if (a.wrapText !== undefined) attrs.push(` wrapText="${a.wrapText ? "1" : "0"}"`);
  if (a.textRotation !== undefined) attrs.push(` textRotation="${a.textRotation}"`);
  if (a.indent !== undefined) attrs.push(` indent="${a.indent}"`);
  if (a.shrinkToFit !== undefined) attrs.push(` shrinkToFit="${a.shrinkToFit ? "1" : "0"}"`);
  if (a.readingOrder !== undefined) attrs.push(` readingOrder="${a.readingOrder}"`);
  if (a.justifyLastLine !== undefined) attrs.push(` justifyLastLine="${a.justifyLastLine ? "1" : "0"}"`);
  if (a.relativeIndent !== undefined) attrs.push(` relativeIndent="${a.relativeIndent}"`);
  return `<alignment${attrs.join("")}/>`;
}

function emitProtection(p: StyleProtection): string {
  const attrs: string[] = [];
  if (p.locked !== undefined) attrs.push(` locked="${p.locked ? "1" : "0"}"`);
  if (p.hidden !== undefined) attrs.push(` hidden="${p.hidden ? "1" : "0"}"`);
  return `<protection${attrs.join("")}/>`;
}

function emitColor(tag: string, c: StyleColor): string {
  const attrs: string[] = [];
  if (c.auto) attrs.push(` auto="1"`);
  if (c.indexed !== undefined) attrs.push(` indexed="${c.indexed}"`);
  if (c.theme !== undefined) attrs.push(` theme="${c.theme}"`);
  if (c.tint !== undefined) attrs.push(` tint="${formatNumber(c.tint)}"`);
  if (c.rgb !== undefined) attrs.push(` rgb="${escapeAttr(c.rgb)}"`);
  return `<${tag}${attrs.join("")}/>`;
}

function emitBoolEl(tag: string, value: boolean): string {
  return value ? `<${tag}/>` : `<${tag} val="0"/>`;
}

function escapeAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function formatNumber(n: number): string {
  if (Number.isInteger(n)) return n.toString();
  return n.toString();
}
