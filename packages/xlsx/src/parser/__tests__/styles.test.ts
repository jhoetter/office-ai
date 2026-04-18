import { describe, expect, it } from "vitest";
import { parseStylesXml } from "../styles.js";
import { serializeStylesXml } from "../../serializer/styles.js";

const SAMPLE_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="0.0000"/>
    <numFmt numFmtId="165" formatCode="&quot;$&quot;#,##0.00"/>
  </numFmts>
  <fonts count="2">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
    <font>
      <b/>
      <sz val="14"/>
      <color rgb="FFFF0000"/>
      <name val="Arial"/>
    </font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFEEAA"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FF000000"/></left>
      <right style="thin"><color rgb="FF000000"/></right>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="164" fontId="0" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>
</styleSheet>`;

describe("parseStylesXml", () => {
  it("parses numFmts with their format codes", () => {
    const t = parseStylesXml(SAMPLE_STYLES_XML);
    expect(t.numFmts.size).toBe(2);
    expect(t.numFmts.get(164)).toBe("0.0000");
    expect(t.numFmts.get(165)).toBe('"$"#,##0.00');
  });

  it("parses fonts including bold + colour + family", () => {
    const t = parseStylesXml(SAMPLE_STYLES_XML);
    expect(t.fonts).toHaveLength(2);
    expect(t.fonts[0].name).toBe("Calibri");
    expect(t.fonts[0].size).toBe(11);
    expect(t.fonts[1].bold).toBe(true);
    expect(t.fonts[1].size).toBe(14);
    expect(t.fonts[1].color?.rgb).toBe("FFFF0000");
    expect(t.fonts[1].name).toBe("Arial");
  });

  it("parses fills (none, gray125, solid with colours)", () => {
    const t = parseStylesXml(SAMPLE_STYLES_XML);
    expect(t.fills).toHaveLength(3);
    expect(t.fills[0].patternType).toBe("none");
    expect(t.fills[1].patternType).toBe("gray125");
    expect(t.fills[2].patternType).toBe("solid");
    expect(t.fills[2].fgColor?.rgb).toBe("FFFFEEAA");
    expect(t.fills[2].bgColor?.indexed).toBe(64);
  });

  it("parses borders", () => {
    const t = parseStylesXml(SAMPLE_STYLES_XML);
    expect(t.borders).toHaveLength(2);
    expect(t.borders[1].left?.style).toBe("thin");
    expect(t.borders[1].left?.color?.rgb).toBe("FF000000");
    expect(t.borders[1].right?.style).toBe("thin");
  });

  it("parses cellXfs with alignment + applyX flags", () => {
    const t = parseStylesXml(SAMPLE_STYLES_XML);
    expect(t.cellXfs).toHaveLength(3);
    expect(t.cellXfs[1].fontId).toBe(1);
    expect(t.cellXfs[1].applyFont).toBe(true);
    expect(t.cellXfs[2].numFmtId).toBe(164);
    expect(t.cellXfs[2].fillId).toBe(2);
    expect(t.cellXfs[2].alignment?.horizontal).toBe("center");
    expect(t.cellXfs[2].alignment?.wrapText).toBe(true);
  });

  it("preserves opaque cellStyles / tableStyles sections", () => {
    const t = parseStylesXml(SAMPLE_STYLES_XML);
    expect(t.cellStylesXml).toContain("Normal");
    expect(t.tableStylesXml).toContain("TableStyleMedium9");
  });
});

describe("serializeStylesXml round-trip", () => {
  it("re-parses to a structurally equivalent table", () => {
    const t1 = parseStylesXml(SAMPLE_STYLES_XML);
    const xml = serializeStylesXml(t1);
    const t2 = parseStylesXml(xml);

    expect(t2.numFmts.size).toBe(t1.numFmts.size);
    for (const [id, code] of t1.numFmts) {
      expect(t2.numFmts.get(id)).toBe(code);
    }
    expect(t2.fonts).toHaveLength(t1.fonts.length);
    expect(t2.fonts[1].bold).toBe(true);
    expect(t2.fonts[1].color?.rgb).toBe("FFFF0000");
    expect(t2.fills).toHaveLength(t1.fills.length);
    expect(t2.fills[2].fgColor?.rgb).toBe("FFFFEEAA");
    expect(t2.borders).toHaveLength(t1.borders.length);
    expect(t2.borders[1].left?.style).toBe("thin");
    expect(t2.cellXfs).toHaveLength(t1.cellXfs.length);
    expect(t2.cellXfs[2].alignment?.horizontal).toBe("center");
    expect(t2.cellXfs[2].applyFill).toBe(true);
    expect(t2.cellStylesXml).not.toBe("");
  });
});
