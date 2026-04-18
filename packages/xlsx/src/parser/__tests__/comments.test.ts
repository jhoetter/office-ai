import { describe, expect, it } from "vitest";
import { parseCommentsPart } from "../comments.js";
import { serializeCommentsPart } from "../../serializer/comments.js";

const SAMPLE_COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>OfficeAI</author>
    <author>Reviewer</author>
  </authors>
  <commentList>
    <comment ref="B7" authorId="0">
      <text>
        <r><t xml:space="preserve">Verify with finance</t></r>
      </text>
    </comment>
    <comment ref="C9" authorId="1">
      <text>
        <r><t>Looks good</t></r>
        <r><t> — please double-check</t></r>
      </text>
    </comment>
  </commentList>
</comments>`;

describe("parseCommentsPart", () => {
  it("extracts authors and comments with stable ids", () => {
    const { authors, comments } = parseCommentsPart(SAMPLE_COMMENTS_XML);
    expect(authors).toEqual(["OfficeAI", "Reviewer"]);
    expect(comments).toHaveLength(2);
    expect(comments[0]).toMatchObject({
      id: "comment-1",
      ref: "B7",
      author: "OfficeAI",
      text: "Verify with finance",
    });
    expect(comments[1]).toMatchObject({
      id: "comment-2",
      ref: "C9",
      author: "Reviewer",
      text: "Looks good — please double-check",
    });
  });

  it("returns empty arrays for an empty comments document", () => {
    const xml = `<?xml version="1.0"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors/><commentList/></comments>`;
    const { authors, comments } = parseCommentsPart(xml);
    expect(authors).toEqual([]);
    expect(comments).toEqual([]);
  });

  it("falls back to empty author when authorId is out of range", () => {
    const xml = `<?xml version="1.0"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>solo</author></authors><commentList><comment ref="A1" authorId="9"><text><r><t>orphan</t></r></text></comment></commentList></comments>`;
    const { comments } = parseCommentsPart(xml);
    expect(comments[0].author).toBe("");
    expect(comments[0].text).toBe("orphan");
  });
});

describe("serializeCommentsPart ↔ parseCommentsPart round-trip", () => {
  it("round-trips a typed comments tuple", () => {
    const { authors, comments } = parseCommentsPart(SAMPLE_COMMENTS_XML);
    const xml = serializeCommentsPart(authors, comments);
    const re = parseCommentsPart(xml);
    expect(re.authors).toEqual(authors);
    expect(re.comments.map((c) => ({ ref: c.ref, author: c.author, text: c.text }))).toEqual(
      comments.map((c) => ({ ref: c.ref, author: c.author, text: c.text }))
    );
  });

  it("escapes XML metacharacters in author and text bodies", () => {
    const xml = serializeCommentsPart(
      ["A & <B>"],
      [{ id: "comment-1", ref: "A1", author: "A & <B>", text: 'quote "x" & <y>' }]
    );
    expect(xml).toContain("A &amp; &lt;B&gt;");
    expect(xml).toContain('quote "x" &amp; &lt;y&gt;');
    const re = parseCommentsPart(xml);
    expect(re.authors[0]).toBe("A & <B>");
    expect(re.comments[0].text).toBe('quote "x" & <y>');
  });

  it('preserves leading/trailing whitespace via xml:space="preserve"', () => {
    const xml = serializeCommentsPart(
      ["a"],
      [{ id: "comment-1", ref: "A1", author: "a", text: "  padded  " }]
    );
    const re = parseCommentsPart(xml);
    expect(re.comments[0].text).toBe("  padded  ");
  });
});
