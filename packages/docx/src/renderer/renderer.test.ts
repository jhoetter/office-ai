import { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { docxSchema } from "./schema.js";
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
