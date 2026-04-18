import { describe, expect, it } from "vitest";
import { DocxAgent } from "../agent/agent.js";
import { chunkIntoPages, type Measure } from "./page-chunker.js";
import { DEFAULT_DOC_ROOT_ATTRS, makeSyntheticDocx } from "../test-utils/synthetic.js";

async function snapshotFromXml(documentXml: string) {
  const buf = await makeSyntheticDocx({ documentXml });
  const agent = await DocxAgent.fromBuffer(buf);
  return agent.getSnapshot();
}

const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>';

function paragraphsXml(count: number): string {
  return Array.from(
    { length: count },
    (_, i) => `<w:p><w:r><w:t xml:space="preserve">para ${i + 1}</w:t></w:r></w:p>`
  ).join("");
}

describe("chunkIntoPages (P3.3 / W9)", () => {
  it("returns one page for a single-paragraph document", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>${paragraphsXml(1)}${SECT_PR}</w:body></w:document>`;
    const snap = await snapshotFromXml(xml);
    const pages = chunkIntoPages(snap);
    expect(pages).toHaveLength(1);
    expect(pages[0].pageNumber).toBe(1);
    expect(pages[0].pageWithinSection).toBe(1);
    expect(pages[0].startBlock).toBe(0);
    expect(pages[0].endBlock).toBe(1);
    expect(pages[0].sectionIndex).toBe(1);
    expect(pages[0].geometry.pgSz.h).toBe(15840);
  });

  it('breaks on an explicit <w:br w:type="page"/> inside a paragraph', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:r><w:t xml:space="preserve">first</w:t></w:r></w:p>
  <w:p><w:r><w:br w:type="page"/><w:t xml:space="preserve">second</w:t></w:r></w:p>
  ${SECT_PR}
</w:body></w:document>`;
    const snap = await snapshotFromXml(xml);
    const pages = chunkIntoPages(snap);
    expect(pages.map((p) => p.pageNumber)).toEqual([1, 2]);
    expect(pages[0].endBlock).toBe(1);
    expect(pages[1].startBlock).toBe(1);
  });

  it("honors <w:lastRenderedPageBreak/> hints when no measure is provided", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:r><w:t xml:space="preserve">first</w:t></w:r></w:p>
  <w:p><w:r><w:lastRenderedPageBreak/><w:t xml:space="preserve">second</w:t></w:r></w:p>
  <w:p><w:r><w:t xml:space="preserve">third</w:t></w:r></w:p>
  ${SECT_PR}
</w:body></w:document>`;
    const snap = await snapshotFromXml(xml);
    const pages = chunkIntoPages(snap);
    expect(pages).toHaveLength(2);
    expect(pages[0].endBlock).toBe(1);
    expect(pages[1].startBlock).toBe(1);
    expect(pages[1].endBlock).toBe(3);
  });

  it("ignores Word's <w:lastRenderedPageBreak/> hint when a measure function is provided", async () => {
    // Word's `<w:lastRenderedPageBreak/>` is computed against Word's
    // own font / line metrics. In the browser our metrics differ
    // (font fallback, line-height, hyphenation), so trusting the hint
    // produced large blank gaps where the chunker flushed earlier
    // than our actually-measured content needed. Once measurement is
    // available, measurement is the source of truth — the saved hint
    // is only consulted on the no-measure code path used by Node-side
    // tests and the very first render frame before the measurement
    // pass settles.
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:r><w:t xml:space="preserve">first</w:t></w:r></w:p>
  <w:p><w:r><w:lastRenderedPageBreak/><w:t xml:space="preserve">second</w:t></w:r></w:p>
  ${SECT_PR}
</w:body></w:document>`;
    const snap = await snapshotFromXml(xml);
    const measure: Measure = () => 10;
    const pages = chunkIntoPages(snap, measure);
    expect(pages.map((p) => p.pageNumber)).toEqual([1]);
    expect(pages[0].endBlock).toBe(2);
  });

  it("honours <w:pageBreakBefore/> on a paragraph property like an explicit hard break (Phase 1)", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:r><w:t xml:space="preserve">first</w:t></w:r></w:p>
  <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t xml:space="preserve">forced second page</w:t></w:r></w:p>
  ${SECT_PR}
</w:body></w:document>`;
    const snap = await snapshotFromXml(xml);
    const pages = chunkIntoPages(snap);
    expect(pages.map((p) => p.pageNumber)).toEqual([1, 2]);
    expect(pages[0].endBlock).toBe(1);
    expect(pages[1].startBlock).toBe(1);
  });

  it("keeps a `keepNext` paragraph together with the next block (Phase 1)", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:r><w:t xml:space="preserve">filler</w:t></w:r></w:p>
  <w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t xml:space="preserve">heading</w:t></w:r></w:p>
  <w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>
  ${SECT_PR}
</w:body></w:document>`;
    const snap = await snapshotFromXml(xml);
    // Content area = 12960 twips. Make the filler 8000 twips, the
    // heading 3000 twips, the body 3000 twips. Without keep-next the
    // chunker would land filler+heading on page 1 (8000+3000 = 11000)
    // and the body on page 2. With keep-next the chunker flushes
    // BEFORE the heading so heading + body land together on page 2.
    const heights = [8000, 3000, 3000];
    const measure: Measure = (i) => heights[i] ?? 0;
    const pages = chunkIntoPages(snap, measure);
    expect(pages).toHaveLength(2);
    expect(pages[0].endBlock).toBe(1); // filler only
    expect(pages[1].startBlock).toBe(1); // heading + body
    expect(pages[1].endBlock).toBe(3);
  });

  it("flushes a page when measured content overflows the section content height", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>${paragraphsXml(5)}${SECT_PR}</w:body></w:document>`;
    const snap = await snapshotFromXml(xml);
    // pgSz.h - top - bottom = 15840 - 1440 - 1440 = 12960 twips of content.
    // Make each block 4000 twips → 3 fit per page (4000 * 3 = 12000 ≤ 12960
    // but the next 4000 would overflow). Five blocks → two pages with 3/2.
    const measure: Measure = () => 4000;
    const pages = chunkIntoPages(snap, measure);
    expect(pages.map((p) => p.pageNumber)).toEqual([1, 2]);
    expect(pages[0].endBlock - pages[0].startBlock).toBe(3);
    expect(pages[1].endBlock - pages[1].startBlock).toBe(2);
  });

  it("never produces an empty leading page when a hard break is the very first signal", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:r><w:br w:type="page"/><w:t xml:space="preserve">solo</w:t></w:r></w:p>
  ${SECT_PR}
</w:body></w:document>`;
    const snap = await snapshotFromXml(xml);
    const pages = chunkIntoPages(snap);
    expect(pages).toHaveLength(1);
    expect(pages[0].startBlock).toBe(0);
  });

  it("starts a new page when crossing a section boundary", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:r><w:t xml:space="preserve">section 1 only para</w:t></w:r></w:p>
  ${SECT_PR}
  <w:p><w:r><w:t xml:space="preserve">section 2 only para</w:t></w:r></w:p>
  ${SECT_PR.replace("12240", "15840").replace("15840", "12240")}
</w:body></w:document>`;
    const snap = await snapshotFromXml(xml);
    const pages = chunkIntoPages(snap);
    expect(pages.map((p) => p.pageNumber)).toEqual([1, 2]);
    // First page picks up section index 1 (the first sectPr).
    expect(pages[0].sectionIndex).toBe(1);
    // Second page picks up section index 3 (the second sectPr).
    expect(pages[1].sectionIndex).toBe(3);
    expect(pages[1].pageWithinSection).toBe(1);
  });

  it("falls back to default geometry when the body has no terminating sectPr", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:p><w:r><w:t xml:space="preserve">no sectPr</w:t></w:r></w:p></w:body></w:document>`;
    const snap = await snapshotFromXml(xml);
    const pages = chunkIntoPages(snap);
    expect(pages).toHaveLength(1);
    expect(pages[0].geometry.pgSz.w).toBe(12240);
    expect(pages[0].geometry.pgMar.top).toBe(1440);
  });

  it("emits a single blank page for an empty body", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body></w:body></w:document>`;
    const snap = await snapshotFromXml(xml);
    const pages = chunkIntoPages(snap);
    expect(pages).toHaveLength(1);
    expect(pages[0].startBlock).toBe(0);
    expect(pages[0].endBlock).toBe(0);
  });
});
