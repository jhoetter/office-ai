import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "./agent.js";
import { snapshotToMarkdown } from "./markdown.js";
import { getPageInfos, getPageMarkdown, pageForParagraph } from "./pages.js";
import { DEFAULT_DOC_ROOT_ATTRS, escapeXml, makeSyntheticDocx } from "../test-utils/synthetic.js";

function syntheticDocXml(paras: ReadonlyArray<string>): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>${paras.join("")}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`;
}

function paraXml(text: string): string {
  return `<w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function paraWithPageBreak(text: string): string {
  // A paragraph that begins with a page-break leaf (so the chunker
  // splits before it).
  return `<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

async function makeAgent(documentXml: string): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

describe("getPageInfos (P3.6 / W22)", () => {
  it("returns a single doc-start page for a one-page document", async () => {
    const agent = await makeAgent(syntheticDocXml([paraXml("Hello world")]));
    const pages = agent.getPages();
    expect(pages).toHaveLength(1);
    expect(pages[0].pageNumber).toBe(1);
    expect(pages[0].trigger).toBe("doc-start");
    expect(pages[0].preview).toContain("Hello world");
  });

  it("splits at a hard <w:br w:type=page/> and labels the trigger", async () => {
    const agent = await makeAgent(
      syntheticDocXml([paraXml("Page one body"), paraWithPageBreak("Page two body")])
    );
    const pages = agent.getPages();
    expect(pages).toHaveLength(2);
    expect(pages[0].trigger).toBe("doc-start");
    expect(pages[1].trigger).toBe("page-break");
    expect(pages[1].preview).toContain("Page two body");
  });

  it("startBlockIndex / endBlockIndex form a half-open range", async () => {
    const agent = await makeAgent(
      syntheticDocXml([paraXml("a"), paraXml("b"), paraWithPageBreak("c"), paraXml("d")])
    );
    const pages = agent.getPages();
    expect(pages[0].startBlockIndex).toBe(0);
    expect(pages[0].endBlockIndex).toBe(2);
    expect(pages[1].startBlockIndex).toBe(2);
    expect(pages[1].endBlockIndex).toBe(4);
  });
});

describe("pageForParagraph (P3.6 / W24)", () => {
  it("resolves the page of every body paragraph", async () => {
    const agent = await makeAgent(
      syntheticDocXml([paraXml("first"), paraXml("second"), paraWithPageBreak("third"), paraXml("fourth")])
    );
    expect(agent.pageForParagraph(0)).toBe(1);
    expect(agent.pageForParagraph(1)).toBe(1);
    expect(agent.pageForParagraph(2)).toBe(2);
    expect(agent.pageForParagraph(3)).toBe(2);
  });

  it("returns null for out-of-range indices", async () => {
    const agent = await makeAgent(syntheticDocXml([paraXml("only")]));
    expect(agent.pageForParagraph(-1)).toBeNull();
    expect(agent.pageForParagraph(99)).toBeNull();
  });

  it("agrees with the snapshot pageForParagraph helper", async () => {
    const agent = await makeAgent(syntheticDocXml([paraXml("a"), paraWithPageBreak("b")]));
    const snap = agent.getSnapshot();
    expect(pageForParagraph(snap, 0)).toBe(1);
    expect(pageForParagraph(snap, 1)).toBe(2);
  });
});

describe("getPageMarkdown (P3.6 / W24)", () => {
  it("returns the markdown for a single page", async () => {
    const agent = await makeAgent(syntheticDocXml([paraXml("Page one"), paraWithPageBreak("Page two")]));
    const md = agent.getPageMarkdown(2);
    expect(md).toContain("Page two");
    expect(md).not.toContain("Page one");
  });

  it("returns null for out-of-range pages", async () => {
    const agent = await makeAgent(syntheticDocXml([paraXml("only")]));
    expect(agent.getPageMarkdown(0)).toBeNull();
    expect(agent.getPageMarkdown(99)).toBeNull();
  });
});

describe("snapshotToMarkdown withPageSections (P3.6 / W22)", () => {
  it("is byte-identical to the default output when withPageSections is false/absent", async () => {
    const agent = await makeAgent(syntheticDocXml([paraXml("Hello"), paraWithPageBreak("World")]));
    const baseline = snapshotToMarkdown(agent.getSnapshot());
    const explicit = snapshotToMarkdown(agent.getSnapshot(), { withPageSections: false });
    expect(explicit).toBe(baseline);
  });

  it("emits page anchors and headings when enabled", async () => {
    const agent = await makeAgent(syntheticDocXml([paraXml("Hello"), paraWithPageBreak("World")]));
    const md = snapshotToMarkdown(agent.getSnapshot(), { withPageSections: true });
    expect(md).toContain("<!-- page 1 -->");
    expect(md).toContain("## Page 1");
    expect(md).toContain("<!-- page 2 -->");
    expect(md).toContain("## Page 2");
    // Page 1 heading appears before page 2 heading.
    expect(md.indexOf("## Page 1")).toBeLessThan(md.indexOf("## Page 2"));
  });

  it("agrees with getPageInfos on the total page count", async () => {
    const agent = await makeAgent(
      syntheticDocXml([paraXml("a"), paraWithPageBreak("b"), paraWithPageBreak("c")])
    );
    const pages = getPageInfos(agent.getSnapshot());
    const md = snapshotToMarkdown(agent.getSnapshot(), { withPageSections: true });
    for (const p of pages) {
      expect(md).toContain(`<!-- page ${p.pageNumber} -->`);
    }
    expect(pages).toHaveLength(3);
  });
});

describe("getPageMarkdown integrates with getPageInfos", () => {
  it("returns markdown matching each page preview", async () => {
    const agent = await makeAgent(syntheticDocXml([paraXml("alpha"), paraWithPageBreak("beta")]));
    const pages = agent.getPages();
    for (const p of pages) {
      const md = getPageMarkdown(agent.getSnapshot(), p.pageNumber);
      expect(md).not.toBeNull();
      expect(md!.includes(p.preview.split(" ")[0])).toBe(true);
    }
  });
});
