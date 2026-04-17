import { describe, expect, it } from "vitest";
import { classifyOpaqueTag, extractOpaqueText } from "./opaque-classification.js";
import type { OpaqueXml } from "./types.js";

function opaque(tag: string, subtree: ReadonlyArray<unknown>): OpaqueXml {
  return { tag, attrs: {}, subtree, rawAttrs: {} };
}

describe("opaque classification", () => {
  it("classifies known structural metadata tags as 'metadata'", () => {
    const metadataTags = [
      "w:bookmarkStart",
      "w:bookmarkEnd",
      "w:fldChar",
      "w:instrText",
      "w:lastRenderedPageBreak",
      "w:proofErr",
      "w:permStart",
      "w:permEnd",
      "w:moveToRangeStart",
    ];
    for (const t of metadataTags) {
      expect(classifyOpaqueTag(t)).toBe("metadata");
    }
  });

  it("classifies known content-wrapper tags as 'content-wrapper'", () => {
    const wrapperTags = ["w:sdt", "w:sdtContent", "w:fldSimple", "mc:AlternateContent", "mc:Fallback"];
    for (const t of wrapperTags) {
      expect(classifyOpaqueTag(t)).toBe("content-wrapper");
    }
  });

  it("falls back to 'placeholder' for any other tag", () => {
    expect(classifyOpaqueTag("w:bizarreUnknown")).toBe("placeholder");
    expect(classifyOpaqueTag("custom:thing")).toBe("placeholder");
    expect(classifyOpaqueTag("")).toBe("placeholder");
  });

  it("extracts text from a flat preserveOrder subtree", () => {
    const o = opaque("w:t", [{ "#text": "hello" }, { "#text": " world" }]);
    expect(extractOpaqueText(o)).toBe("hello world");
  });

  it("extracts text from a deeply nested SDT-like subtree", () => {
    const o = opaque("w:sdt", [
      { "w:sdtPr": [{ "w:alias": [], ":@": { "@_w:val": "TOC" } }] },
      {
        "w:sdtContent": [
          { "w:p": [{ "w:r": [{ "w:t": [{ "#text": "TOC heading" }] }] }] },
          {
            "w:p": [
              { "w:r": [{ "w:t": [{ "#text": "1 Introduction" }] }] },
              { "w:r": [{ "w:t": [{ "#text": " ............... 1" }] }] },
            ],
          },
        ],
      },
    ]);
    const text = extractOpaqueText(o);
    expect(text).toContain("TOC heading");
    expect(text).toContain("1 Introduction");
    expect(text).toContain("............... 1");
  });

  it("returns empty string for subtrees with no #text descendants", () => {
    const o = opaque("w:bookmarkStart", []);
    expect(extractOpaqueText(o)).toBe("");
  });

  it("skips text inside nested metadata-classified children", () => {
    // A `<w:sdt>` wrapping a TOC field looks like
    //   <w:sdt>
    //     <w:sdtContent>
    //       <w:r><w:instrText> TOC \h \o "1-3" </w:instrText></w:r>
    //       <w:r><w:t>Inhaltsverzeichnis</w:t></w:r>
    //     </w:sdtContent>
    //   </w:sdt>
    // The field instruction is runtime markup, not user-visible content,
    // and must not leak into the wrapper preview text.
    const o = opaque("w:sdt", [
      {
        "w:sdtContent": [
          {
            "w:r": [{ "w:instrText": [{ "#text": ' TOC \\h \\o "1-3" ' }] }],
          },
          {
            "w:r": [{ "w:t": [{ "#text": "Inhaltsverzeichnis" }] }],
          },
        ],
      },
    ]);
    const text = extractOpaqueText(o);
    expect(text).toBe("Inhaltsverzeichnis");
    expect(text).not.toContain("TOC");
  });

  it("ignores attribute maps and processing instructions when walking", () => {
    const o = opaque("w:r", [
      { ":@": { "@_w:rsidR": "0001" } },
      { "?xml": [], ":@": { "@_version": "1.0" } },
      { "w:t": [{ "#text": "actual text" }] },
    ]);
    expect(extractOpaqueText(o)).toBe("actual text");
  });
});
