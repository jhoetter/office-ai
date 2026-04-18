import { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { docxSchema, paragraphHtmlTag } from "./schema.js";
import { docToPM } from "./doc-to-pm.js";
import { transactionToCommands } from "./transaction-to-commands.js";
import { DocxAgent } from "../agent/agent.js";
import { makeSyntheticDocx, plainDocxXml } from "../test-utils/synthetic.js";

async function loadAgent(paragraphs: { text: string; styleId?: string }[]): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: plainDocxXml(paragraphs) });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function stateFor(agent: DocxAgent): EditorState {
  return EditorState.create({ schema: docxSchema, doc: docToPM(agent.getSnapshot()) });
}

describe("renderer", () => {
  it("docToPM projects paragraphs and runs", async () => {
    const agent = await loadAgent([{ text: "Hello", styleId: "Heading1" }, { text: "world" }]);
    const pm = docToPM(agent.getSnapshot());
    expect(pm.type.name).toBe("doc");
    const paragraphs: import("prosemirror-model").Node[] = [];
    pm.forEach((c) => {
      if (c.type.name === "paragraph") paragraphs.push(c);
    });
    expect(paragraphs).toHaveLength(2);
    const p0 = paragraphs[0]!;
    expect(p0.attrs.styleId).toBe("Heading1");
    expect(p0.textContent).toBe("Hello");
    expect(paragraphs[1]!.textContent).toBe("world");
  });

  it("docToPM applies bold/italic marks from run properties", async () => {
    const agent = await loadAgent([{ text: "alpha beta" }]);
    await agent.applyCommand({
      type: "docx:format-range",
      payload: {
        range: {
          start: { paragraph: 0, run: 0, offset: 0 },
          end: { paragraph: 0, run: 0, offset: 5 },
        },
        format: { bold: true },
      },
      source: "human",
    });
    const pm = docToPM(agent.getSnapshot());
    const para = pm.child(0);
    let sawBold = false;
    para.descendants((n) => {
      if (n.isText && n.marks.some((m) => m.type.name === "bold")) sawBold = true;
      return true;
    });
    expect(sawBold).toBe(true);
  });

  it("typing inserts a docx:insert-text command at the right position", async () => {
    const agent = await loadAgent([{ text: "Hello" }]);
    const state = stateFor(agent);
    const insertAt = 1; // inside paragraph, before "H"
    const tx = state.tr.insertText("X", insertAt, insertAt);
    const result = transactionToCommands(tx, state);
    expect(result.unsupported).toHaveLength(0);
    expect(result.commands).toHaveLength(1);
    const cmd = result.commands[0];
    expect(cmd.type).toBe("docx:insert-text");
    expect(cmd.payload).toMatchObject({ at: { paragraph: 0, offset: 0 }, text: "X" });
  });

  it("selection delete produces a docx:delete-range command", async () => {
    const agent = await loadAgent([{ text: "Hello world" }]);
    const state = stateFor(agent);
    const tx = state.tr.delete(7, 12); // "world" inside "Hello world"
    const result = transactionToCommands(tx, state);
    expect(result.commands).toHaveLength(1);
    const cmd = result.commands[0];
    expect(cmd.type).toBe("docx:delete-range");
    expect(cmd.payload).toMatchObject({
      range: {
        start: { paragraph: 0, offset: 6 },
        end: { paragraph: 0, offset: 11 },
      },
    });
  });

  it("addMark over a range produces docx:format-range", async () => {
    const agent = await loadAgent([{ text: "alpha beta" }]);
    const state = stateFor(agent);
    const tx = state.tr.addMark(1, 6, docxSchema.marks.bold.create());
    const result = transactionToCommands(tx, state);
    expect(result.commands).toHaveLength(1);
    const cmd = result.commands[0];
    expect(cmd.type).toBe("docx:format-range");
    expect(cmd.payload).toMatchObject({
      range: { start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 5 } },
      format: { bold: true },
    });
  });

  it("after dispatching commands through the agent, docToPM reflects new text", async () => {
    const agent = await loadAgent([{ text: "Hello" }]);
    const state = stateFor(agent);
    const tx = state.tr.insertText("X", 1, 1);
    const result = transactionToCommands(tx, state);
    await agent.applyCommands(result.commands);
    const pm = docToPM(agent.getSnapshot());
    expect(pm.child(0).textContent).toBe("XHello");
  });

  describe("paragraph toDOM emits heading tags & alignment", () => {
    it("paragraphHtmlTag maps English & German heading style ids", () => {
      expect(paragraphHtmlTag("Heading1")).toBe("h1");
      expect(paragraphHtmlTag("Heading2")).toBe("h2");
      expect(paragraphHtmlTag("Heading3")).toBe("h3");
      expect(paragraphHtmlTag("Heading4")).toBe("h4");
      expect(paragraphHtmlTag("Heading5")).toBe("h5");
      expect(paragraphHtmlTag("Heading6")).toBe("h6");
      expect(paragraphHtmlTag("Heading7")).toBe("h6");
      expect(paragraphHtmlTag("Title")).toBe("h1");
      expect(paragraphHtmlTag("Subtitle")).toBe("h2");
      // German Word strips the leading 'Ü' from `Überschrift` because
      // OOXML style ids must be ASCII.
      expect(paragraphHtmlTag("berschrift1")).toBe("h1");
      expect(paragraphHtmlTag("berschrift3")).toBe("h3");
      expect(paragraphHtmlTag("Untertitel")).toBe("h2");
      expect(paragraphHtmlTag("Titel")).toBe("h1");
      // Unknown / body styles fall through to <p>.
      expect(paragraphHtmlTag("")).toBe("p");
      expect(paragraphHtmlTag("Normal")).toBe("p");
      expect(paragraphHtmlTag("BodyText")).toBe("p");
    });

    it("paragraph.toDOM emits the matching heading tag for known styleIds", () => {
      const node = docxSchema.nodes.paragraph.create({ styleId: "Heading2" });
      const dom = node.type.spec.toDOM!(node) as [string, Record<string, string>, number];
      expect(dom[0]).toBe("h2");
      expect(dom[1]["data-style"]).toBe("Heading2");
    });

    it("paragraph.toDOM emits style=text-align when alignment is set", () => {
      const node = docxSchema.nodes.paragraph.create({ alignment: "center" });
      const dom = node.type.spec.toDOM!(node) as [string, Record<string, string>, number];
      expect(dom[0]).toBe("p");
      expect(dom[1].style).toBe("text-align:center");
      expect(dom[1]["data-align"]).toBe("center");
    });

    it("paragraph.toDOM does not emit a style attr when alignment is null", () => {
      const node = docxSchema.nodes.paragraph.create({});
      const dom = node.type.spec.toDOM!(node) as [string, Record<string, string>, number];
      expect("style" in dom[1]).toBe(false);
      expect("data-align" in dom[1]).toBe(false);
    });
  });

  describe("image rendering — real <img> from media (P2.4 / W22)", () => {
    /**
     * The dataUrl/width/height/alt attrs are populated by docToPM via the
     * MediaResolver. We exercise the full pipeline against the real-world
     * 05-inline-image fixture so a regression in either the parser, the
     * resolver, or the schema's toDOM mapping is caught.
     */
    it("docToPM populates dataUrl + intrinsic dimensions for inline images", async () => {
      const fs = await import("node:fs/promises");
      const path = await import("node:path");
      const buf = await fs.readFile(
        path.resolve(__dirname, "../../../../fixtures/docx/real-world/05-inline-image.docx")
      );
      const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
      const pm = docToPM(agent.getSnapshot());
      let imageNode: import("prosemirror-model").Node | null = null;
      pm.descendants((n) => {
        if (n.type.name === "image") {
          imageNode = n;
          return false;
        }
        return true;
      });
      expect(imageNode).not.toBeNull();
      if (!imageNode) return;
      const node = imageNode as import("prosemirror-model").Node;
      const dataUrl = node.attrs.dataUrl as string | null;
      expect(typeof dataUrl).toBe("string");
      expect(dataUrl).toMatch(/^data:image\/png;base64,/);
      // 228600 EMU @ 9525 EMU/px = 24px on each side for the fixture image.
      expect(node.attrs.width).toBe(24);
      expect(node.attrs.height).toBe(24);
    });

    it("image.toDOM emits <img src=data:...> when dataUrl is present", () => {
      const node = docxSchema.nodes.image.create({
        dataUrl: "data:image/png;base64,iVBORw0KGgo=",
        width: 32,
        height: 32,
        alt: "smoke",
      });
      const dom = node.type.spec.toDOM!(node) as [string, Record<string, string>];
      expect(dom[0]).toBe("img");
      expect(dom[1].src).toBe("data:image/png;base64,iVBORw0KGgo=");
      expect(dom[1].width).toBe("32");
      expect(dom[1].height).toBe("32");
      expect(dom[1].alt).toBe("smoke");
      expect(dom[1].class).toBe("pm-image");
    });

    it("image.toDOM falls back to [image] placeholder when dataUrl is missing", () => {
      const node = docxSchema.nodes.image.create({});
      const dom = node.type.spec.toDOM!(node) as [string, Record<string, string>, string];
      expect(dom[0]).toBe("span");
      expect(dom[1].class).toBe("pm-image-placeholder");
      expect(dom[2]).toBe("[image]");
    });
  });

  it("agent.subscribe fires on every applied command (single-funnel hook)", async () => {
    const agent = await loadAgent([{ text: "abc" }]);
    let count = 0;
    const off = agent.subscribe(() => {
      count++;
    });
    await agent.applyCommand({
      type: "docx:insert-text",
      payload: { at: { paragraph: 0, run: 0, offset: 0 }, text: "Z" },
      source: "human",
    });
    expect(count).toBe(1);
    off();
  });
});
