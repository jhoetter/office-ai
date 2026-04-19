import { describe, expect, it } from "vitest";
import type {
  BlockNode,
  DocxSnapshot,
  InlineNode,
  Paragraph,
  Run,
  RunChild,
} from "@officeai/docx";
import { docxToMarkdown, docxToText } from "./export-text";

/* ── tiny snapshot factories ─────────────────────────────────────── */

function run(
  text: string,
  props: { bold?: boolean; italic?: boolean; strike?: boolean } = {}
): Run {
  return {
    kind: "run",
    properties: props,
    children: [{ kind: "text", text }],
  } as unknown as Run;
}

function runWithChildren(
  children: ReadonlyArray<RunChild>,
  props: { bold?: boolean; italic?: boolean; strike?: boolean } = {}
): Run {
  return {
    kind: "run",
    properties: props,
    children,
  } as unknown as Run;
}

function paragraph(
  children: ReadonlyArray<InlineNode>,
  properties: Paragraph["properties"] = {}
): Paragraph {
  return {
    kind: "paragraph",
    properties,
    children,
  } as unknown as Paragraph;
}

function snapshot(body: ReadonlyArray<BlockNode>): DocxSnapshot {
  return { root: { body } } as unknown as DocxSnapshot;
}

/* ── docxToText ──────────────────────────────────────────────────── */

describe("docxToText", () => {
  it("serializes a single paragraph as one line", () => {
    const snap = snapshot([paragraph([run("hello world")])]);
    expect(docxToText(snap)).toBe("hello world\n");
  });

  it("joins paragraphs with newlines", () => {
    const snap = snapshot([
      paragraph([run("first")]),
      paragraph([run("second")]),
    ]);
    expect(docxToText(snap)).toBe("first\nsecond\n");
  });

  it("collapses runs of 3+ newlines to a single blank line", () => {
    const snap = snapshot([
      paragraph([run("a")]),
      paragraph([]),
      paragraph([]),
      paragraph([]),
      paragraph([run("b")]),
    ]);
    // Empty paragraphs each emit "" so the join produces multiple "\n".
    // The serializer caps that at \n\n.
    expect(docxToText(snap)).toBe("a\n\nb\n");
  });

  it("expands tabs and breaks inside a run", () => {
    const snap = snapshot([
      paragraph([
        runWithChildren([
          { kind: "text", text: "a" },
          { kind: "tab" },
          { kind: "text", text: "b" },
          { kind: "break" },
          { kind: "text", text: "c" },
        ] as unknown as RunChild[]),
      ]),
    ]);
    expect(docxToText(snap)).toBe("a\tb\nc\n");
  });

  it("renders bullet list paragraphs with a leading dash", () => {
    const snap = snapshot([
      paragraph([run("item 1")], { numbering: { numId: 1, ilvl: 0 } } as Paragraph["properties"]),
      paragraph([run("item 2")], { numbering: { numId: 1, ilvl: 0 } } as Paragraph["properties"]),
    ]);
    expect(docxToText(snap)).toBe("- item 1\n- item 2\n");
  });

  it("indents nested list items", () => {
    const snap = snapshot([
      paragraph([run("parent")], { numbering: { numId: 1, ilvl: 0 } } as Paragraph["properties"]),
      paragraph([run("child")], { numbering: { numId: 1, ilvl: 1 } } as Paragraph["properties"]),
    ]);
    expect(docxToText(snap)).toBe("- parent\n  - child\n");
  });

  it("ignores deletion-only text inside runs", () => {
    const snap = snapshot([
      paragraph([
        runWithChildren([
          { kind: "text", text: "kept" } as unknown as RunChild,
          { kind: "text", text: "gone", isDelText: true } as unknown as RunChild,
        ]),
      ]),
    ]);
    expect(docxToText(snap)).toBe("kept\n");
  });
});

/* ── docxToMarkdown ──────────────────────────────────────────────── */

describe("docxToMarkdown", () => {
  it("emits H1-H6 prefixes for Heading style ids", () => {
    const snap = snapshot([
      paragraph([run("Title")], { styleId: "Heading1" } as Paragraph["properties"]),
      paragraph([run("Sub")], { styleId: "Heading2" } as Paragraph["properties"]),
      paragraph([run("Body")]),
    ]);
    expect(docxToMarkdown(snap)).toBe("# Title\n\n## Sub\n\nBody\n");
  });

  it("clamps heading levels above 6", () => {
    const snap = snapshot([
      paragraph([run("Deep")], { styleId: "Heading9" } as Paragraph["properties"]),
    ]);
    expect(docxToMarkdown(snap)).toBe("###### Deep\n");
  });

  it("wraps bold/italic/strike runs with the right delimiters", () => {
    const snap = snapshot([
      paragraph([
        run("normal "),
        run("bold", { bold: true }),
        run(" "),
        run("italic", { italic: true }),
        run(" "),
        run("strike", { strike: true }),
      ]),
    ]);
    expect(docxToMarkdown(snap)).toBe("normal **bold** _italic_ ~~strike~~\n");
  });

  it("escapes Markdown special characters in plain runs", () => {
    const snap = snapshot([paragraph([run("a*b_c`d[e]")])]);
    expect(docxToMarkdown(snap)).toBe("a\\*b\\_c\\`d\\[e\\]\n");
  });

  it("renders bullet lists with `-` markers", () => {
    const snap = snapshot([
      paragraph([run("one")], { numbering: { numId: 1, ilvl: 0 } } as Paragraph["properties"]),
      paragraph([run("two")], { numbering: { numId: 1, ilvl: 0 } } as Paragraph["properties"]),
    ]);
    expect(docxToMarkdown(snap)).toBe("- one\n- two\n");
  });

  it("indents nested bullets two spaces per level", () => {
    const snap = snapshot([
      paragraph([run("parent")], { numbering: { numId: 1, ilvl: 0 } } as Paragraph["properties"]),
      paragraph([run("child")], { numbering: { numId: 1, ilvl: 1 } } as Paragraph["properties"]),
    ]);
    expect(docxToMarkdown(snap)).toBe("- parent\n  - child\n");
  });
});
