import { describe, expect, it } from "vitest";
import { DocxAgent } from "../agent/agent.js";
import { resolveEffectiveRpr, WORD_DEFAULT_THEME_FONTS } from "../agent/style-resolver.js";
import { DEFAULT_DOC_ROOT_ATTRS, makeSyntheticDocx } from "../test-utils/synthetic.js";

const SECT_PR = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;

const DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading text</w:t></w:r></w:p>
  <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
${SECT_PR}
</w:body></w:document>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>
      <w:b/>
      <w:sz w:val="36"/>
    </w:rPr>
  </w:style>
</w:styles>`;

const APTOS_THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
  <a:themeElements>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Aptos Display" panose="02110004020202020204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Aptos" panose="02110004020202020204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`;

const CAMBRIA_THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office 2007">
  <a:themeElements>
    <a:fontScheme name="Office 2007">
      <a:majorFont>
        <a:latin typeface="Cambria"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
      </a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`;

describe("ThemePart parser + resolver projection (P3.9)", () => {
  it("parses majorFont.latin / minorFont.latin from theme1.xml", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: DOC_XML,
      extra: {
        "word/styles.xml": STYLES_XML,
        "word/theme/theme1.xml": APTOS_THEME_XML,
      },
    });
    const agent = await DocxAgent.fromBuffer(buf);
    const theme = agent.getSnapshot().root.theme;
    expect(theme).toBeDefined();
    expect(theme?.majorFont.latin).toBe("Aptos Display");
    expect(theme?.minorFont.latin).toBe("Aptos");
  });

  it("captures asciiTheme attributes into typed RunProperties fields", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: DOC_XML,
      extra: {
        "word/styles.xml": STYLES_XML,
      },
    });
    const agent = await DocxAgent.fromBuffer(buf);
    const styles = agent.getSnapshot().root.styles;
    const heading1 = styles?.styles.get("Heading1");
    expect(heading1?.rPr?.fontFamilyAsciiTheme).toBe("majorHAnsi");
    expect(heading1?.rPr?.fontFamilyHAnsiTheme).toBe("majorHAnsi");
    expect(heading1?.rPr?.fontFamily).toBeUndefined();
  });

  it("resolver projects asciiTheme through theme1.xml when present", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: DOC_XML,
      extra: {
        "word/styles.xml": STYLES_XML,
        "word/theme/theme1.xml": APTOS_THEME_XML,
      },
    });
    const agent = await DocxAgent.fromBuffer(buf);
    const snapshot = agent.getSnapshot();
    const heading = resolveEffectiveRpr(snapshot, 0, 0);
    expect(heading.fontFamily).toBe("Aptos Display");
    // Body paragraph has no styleId → falls back to docDefaults Calibri.
    const body = resolveEffectiveRpr(snapshot, 1, 0);
    expect(body.fontFamily).toBe("Calibri");
  });

  it("resolver respects custom theme fonts (Cambria/Calibri)", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: DOC_XML,
      extra: {
        "word/styles.xml": STYLES_XML,
        "word/theme/theme1.xml": CAMBRIA_THEME_XML,
      },
    });
    const agent = await DocxAgent.fromBuffer(buf);
    const snapshot = agent.getSnapshot();
    const heading = resolveEffectiveRpr(snapshot, 0, 0);
    expect(heading.fontFamily).toBe("Cambria");
  });

  it("resolver falls back to Word-default theme map when theme part is absent", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: DOC_XML,
      extra: {
        "word/styles.xml": STYLES_XML,
        // intentionally no theme1.xml — exercises the fallback path
      },
    });
    const agent = await DocxAgent.fromBuffer(buf);
    const snapshot = agent.getSnapshot();
    expect(snapshot.root.theme).toBeUndefined();
    const heading = resolveEffectiveRpr(snapshot, 0, 0);
    expect(heading.fontFamily).toBe(WORD_DEFAULT_THEME_FONTS["majorHAnsi"]);
    expect(heading.fontFamily).toBe("Aptos Display");
  });

  it("explicit literal w:ascii on the run wins over inherited theme ref", async () => {
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Comic Sans MS" w:hAnsi="Comic Sans MS"/></w:rPr><w:t>Override</w:t></w:r></w:p>
${SECT_PR}
</w:body></w:document>`;
    const buf = await makeSyntheticDocx({
      documentXml: docXml,
      extra: {
        "word/styles.xml": STYLES_XML,
        "word/theme/theme1.xml": APTOS_THEME_XML,
      },
    });
    const agent = await DocxAgent.fromBuffer(buf);
    const snapshot = agent.getSnapshot();
    const run = resolveEffectiveRpr(snapshot, 0, 0);
    expect(run.fontFamily).toBe("Comic Sans MS");
  });

  it("Heading1 style replaces docDefaults rFonts wholesale (Word semantics)", async () => {
    // docDefaults has w:ascii="Calibri", Heading1 only has asciiTheme.
    // Per Word, the style replaces the entire <w:rFonts> element, so
    // the literal Calibri is dropped and the resolver must project the
    // theme ref. Without this rule, our resolver would report Calibri
    // and the editor would disagree with what Word renders.
    const buf = await makeSyntheticDocx({
      documentXml: DOC_XML,
      extra: {
        "word/styles.xml": STYLES_XML,
        "word/theme/theme1.xml": APTOS_THEME_XML,
      },
    });
    const agent = await DocxAgent.fromBuffer(buf);
    const heading = resolveEffectiveRpr(agent.getSnapshot(), 0, 0);
    expect(heading.fontFamily).not.toBe("Calibri");
    expect(heading.fontFamily).toBe("Aptos Display");
  });

  it("opaque <w:rFonts> with theme refs round-trips byte-identical when nothing else mutates the run", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: DOC_XML,
      extra: {
        "word/styles.xml": STYLES_XML,
        "word/theme/theme1.xml": APTOS_THEME_XML,
      },
    });
    const agent = await DocxAgent.fromBuffer(buf);
    const exported = await agent.exportFile();
    // Re-parse the export and confirm the heading style still carries
    // the asciiTheme attribute (i.e. the typed projection didn't
    // *replace* the opaque rFonts on disk).
    const second = await DocxAgent.fromBuffer(exported);
    const heading1 = second.getSnapshot().root.styles?.styles.get("Heading1");
    expect(heading1?.rPr?.fontFamilyAsciiTheme).toBe("majorHAnsi");
  });
});
