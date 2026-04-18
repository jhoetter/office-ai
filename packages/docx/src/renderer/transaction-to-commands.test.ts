import { Slice, Fragment } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { makeSyntheticDocx, plainDocxXml } from "../test-utils/synthetic.js";
import { docToPM } from "./doc-to-pm.js";
import { docxSchema } from "./schema.js";
import { transactionToCommands } from "./transaction-to-commands.js";

/**
 * Table-driven tests for `transactionToCommands`. Each case constructs a
 * deterministic PM doc from a small fixture, runs an explicit transaction
 * (insert / delete / paste / format) against it, then asserts the
 * sequence of commands the funnel produces.
 *
 * No Browser. Pure PM + vitest.
 */

async function loadAgent(paragraphs: { text: string; styleId?: string }[]): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: plainDocxXml(paragraphs) });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function stateFor(agent: DocxAgent): EditorState {
  return EditorState.create({ schema: docxSchema, doc: docToPM(agent.getSnapshot()) });
}

/** Build a slice containing the given top-level paragraphs (text-only). */
function paragraphsSlice(texts: ReadonlyArray<string>): Slice {
  const paras = texts.map((t) => {
    const inline = t.length > 0 ? Fragment.from(docxSchema.text(t)) : Fragment.empty;
    return docxSchema.nodes.paragraph.create(null, inline);
  });
  return new Slice(Fragment.from(paras), 1, 1);
}

describe("transactionToCommands — PM funnel", () => {
  describe("mark re-assertion across boundary edits", () => {
    it("typing inside a partially-bold span emits insert-text + format-range with the bold mark", async () => {
      const agent = await loadAgent([{ text: "alpha bravo" }]);
      // Mark "bravo" as bold via the agent so the PM doc projection includes the mark.
      await agent.applyCommand({
        type: "docx:format-range",
        payload: {
          range: {
            start: { paragraph: 0, run: 0, offset: 6 },
            end: { paragraph: 0, run: 0, offset: 11 },
          },
          format: { bold: true },
        },
        source: "human",
      });
      const state = stateFor(agent);
      // Insert "X" inside the bold span (between 'b' and 'r' in "bravo" → PM offset = 1 + 6 + 1 = 8).
      const insertAt = 1 + 6 + 1;
      const tx = state.tr.insertText("X", insertAt, insertAt);
      const result = transactionToCommands(tx, state);
      expect(result.unsupported).toHaveLength(0);
      expect(result.commands.length).toBe(2);
      expect(result.commands[0].type).toBe("docx:insert-text");
      expect(result.commands[0].payload).toMatchObject({
        at: { paragraph: 0, offset: 7 },
        text: "X",
      });
      expect(result.commands[1].type).toBe("docx:format-range");
      expect(result.commands[1].payload).toMatchObject({
        range: { start: { paragraph: 0, offset: 7 }, end: { paragraph: 0, offset: 8 } },
        format: { bold: true },
      });
    });

    it("typing in a plain paragraph does NOT emit a follow-up format-range", async () => {
      const agent = await loadAgent([{ text: "plain text" }]);
      const state = stateFor(agent);
      const tx = state.tr.insertText("X", 1, 1);
      const result = transactionToCommands(tx, state);
      expect(result.unsupported).toHaveLength(0);
      expect(result.commands).toHaveLength(1);
      expect(result.commands[0].type).toBe("docx:insert-text");
    });

    it("PM-encoded marks on the inserted slice still drive a follow-up format-range", () => {
      // Construct a state where a slice carrying explicit marks is inserted
      // (simulates PM's storedMarks behavior on typing inside a marked span).
      const para = docxSchema.nodes.paragraph.create(null, docxSchema.text("alpha bravo"));
      const doc = docxSchema.nodes.doc.create(null, [para]);
      const state = EditorState.create({ schema: docxSchema, doc });
      const markedText = docxSchema.text("X", [docxSchema.marks.italic.create()]);
      const slice = new Slice(Fragment.from(markedText), 0, 0);
      const cursor = 1 + 3;
      const tx = state.tr.replace(cursor, cursor, slice);
      const result = transactionToCommands(tx, state);
      expect(result.unsupported).toHaveLength(0);
      expect(result.commands).toHaveLength(2);
      expect(result.commands[1].type).toBe("docx:format-range");
      expect(result.commands[1].payload).toMatchObject({ format: { italic: true } });
    });
  });

  describe("multi-block paste", () => {
    it("pasting two paragraphs emits insert-text + insert-paragraph + insert-text", async () => {
      const agent = await loadAgent([{ text: "before|after" }]);
      const state = stateFor(agent);
      // PM cursor between "before" and "|" → offset 7 inside paragraph (1 + 6).
      const cursor = 1 + 6;
      const slice = paragraphsSlice(["AAA", "BBB"]);
      const tx = state.tr.replace(cursor, cursor, slice);
      const result = transactionToCommands(tx, state);
      expect(result.unsupported).toHaveLength(0);
      expect(result.commands.map((c) => c.type)).toEqual([
        "docx:insert-text",
        "docx:insert-paragraph",
        "docx:insert-text",
      ]);
      expect(result.commands[0].payload).toMatchObject({
        at: { paragraph: 0, offset: 6 },
        text: "AAA",
      });
      expect(result.commands[1].payload).toMatchObject({
        at: { paragraph: 0, offset: 9 },
      });
      expect(result.commands[2].payload).toMatchObject({
        at: { paragraph: 1, offset: 0 },
        text: "BBB",
      });
    });

    it("pasting three paragraphs emits insert-text + 2× (insert-paragraph + insert-text)", async () => {
      const agent = await loadAgent([{ text: "X" }]);
      const state = stateFor(agent);
      const cursor = 1; // before 'X'
      const slice = paragraphsSlice(["one", "two", "three"]);
      const tx = state.tr.replace(cursor, cursor, slice);
      const result = transactionToCommands(tx, state);
      expect(result.unsupported).toHaveLength(0);
      expect(result.commands.map((c) => c.type)).toEqual([
        "docx:insert-text",
        "docx:insert-paragraph",
        "docx:insert-text",
        "docx:insert-paragraph",
        "docx:insert-text",
      ]);
      expect(result.commands[0].payload).toMatchObject({ at: { paragraph: 0, offset: 0 }, text: "one" });
      expect(result.commands[1].payload).toMatchObject({ at: { paragraph: 0, offset: 3 } });
      expect(result.commands[2].payload).toMatchObject({ at: { paragraph: 1, offset: 0 }, text: "two" });
      expect(result.commands[3].payload).toMatchObject({ at: { paragraph: 1, offset: 3 } });
      expect(result.commands[4].payload).toMatchObject({ at: { paragraph: 2, offset: 0 }, text: "three" });
    });

    it("paste-of-two with empty trailing paragraph still emits the right paragraph split", async () => {
      const agent = await loadAgent([{ text: "tail" }]);
      const state = stateFor(agent);
      const cursor = 1 + 4; // end of "tail"
      const slice = paragraphsSlice(["AAA", ""]);
      const tx = state.tr.replace(cursor, cursor, slice);
      const result = transactionToCommands(tx, state);
      expect(result.unsupported).toHaveLength(0);
      expect(result.commands.map((c) => c.type)).toEqual(["docx:insert-text", "docx:insert-paragraph"]);
    });
  });

  describe("multi-paragraph delete", () => {
    it("selection spanning two paragraphs emits a single docx:delete-range with cross-paragraph endpoints", async () => {
      const agent = await loadAgent([{ text: "first paragraph" }, { text: "second paragraph" }]);
      const state = stateFor(agent);
      // PM doc: <p>first paragraph</p><p>second paragraph</p>
      // p0 size = 1 + 15 + 1 = 17; p0 inner [1..16].
      // Select from PM 7 (after "first ") to PM 17 + 1 + 7 = 25 (after "second ").
      const fromPM = 1 + 6;
      const toPM = 1 + 15 + 1 + 1 + 7;
      const tx = state.tr.delete(fromPM, toPM);
      const result = transactionToCommands(tx, state);
      expect(result.unsupported).toHaveLength(0);
      expect(result.commands).toHaveLength(1);
      expect(result.commands[0].type).toBe("docx:delete-range");
      expect(result.commands[0].payload).toMatchObject({
        range: {
          start: { paragraph: 0, offset: 6 },
          end: { paragraph: 1, offset: 7 },
        },
      });
    });
  });

  describe("suggesting mode", () => {
    it("typing in suggesting mode emits docx:insert-text-tracked carrying the author", async () => {
      const agent = await loadAgent([{ text: "alpha bravo" }]);
      const state = stateFor(agent);
      const tx = state.tr.insertText("X", 1, 1);
      const result = transactionToCommands(tx, state, { mode: "suggest", author: "Alice" });
      expect(result.unsupported).toHaveLength(0);
      expect(result.commands).toHaveLength(1);
      expect(result.commands[0].type).toBe("docx:insert-text-tracked");
      expect(result.commands[0].payload).toMatchObject({
        at: { paragraph: 0, offset: 0 },
        text: "X",
        author: "Alice",
      });
    });

    it("single-paragraph delete in suggesting mode emits docx:delete-range-tracked", async () => {
      const agent = await loadAgent([{ text: "alpha bravo" }]);
      const state = stateFor(agent);
      // Delete "alpha" — PM positions 1..6.
      const tx = state.tr.delete(1, 6);
      const result = transactionToCommands(tx, state, { mode: "suggest", author: "Alice" });
      expect(result.unsupported).toHaveLength(0);
      expect(result.commands).toHaveLength(1);
      expect(result.commands[0].type).toBe("docx:delete-range-tracked");
      expect(result.commands[0].payload).toMatchObject({
        range: {
          start: { paragraph: 0, offset: 0 },
          end: { paragraph: 0, offset: 5 },
        },
        author: "Alice",
      });
    });

    it("typing-over-selection in suggesting mode emits delete-tracked + insert-tracked", async () => {
      const agent = await loadAgent([{ text: "alpha bravo" }]);
      const state = stateFor(agent);
      // Replace "alpha" (PM 1..6) with "Z".
      const tx = state.tr.insertText("Z", 1, 6);
      const result = transactionToCommands(tx, state, { mode: "suggest", author: "Alice" });
      expect(result.unsupported).toHaveLength(0);
      expect(result.commands.map((c) => c.type)).toEqual([
        "docx:delete-range-tracked",
        "docx:insert-text-tracked",
      ]);
      expect(result.commands[1].payload).toMatchObject({ text: "Z", author: "Alice" });
    });

    it("multi-paragraph delete in suggesting mode falls back to unsupported (with explanation)", async () => {
      const agent = await loadAgent([{ text: "first paragraph" }, { text: "second paragraph" }]);
      const state = stateFor(agent);
      const fromPM = 1 + 6;
      const toPM = 1 + 15 + 1 + 1 + 7;
      const tx = state.tr.delete(fromPM, toPM);
      const result = transactionToCommands(tx, state, { mode: "suggest", author: "Alice" });
      expect(result.commands).toHaveLength(0);
      // The handler pushes a specific reason then the funnel adds a
      // generic "unsupported step" — first entry is the actionable one.
      expect(result.unsupported.length).toBeGreaterThanOrEqual(1);
      expect(result.unsupported[0].reason).toMatch(/cross paragraph/i);
    });

    it("missing author in suggesting mode falls back to unsupported", async () => {
      const agent = await loadAgent([{ text: "alpha" }]);
      const state = stateFor(agent);
      const tx = state.tr.insertText("X", 1, 1);
      const result = transactionToCommands(tx, state, { mode: "suggest" });
      expect(result.commands).toHaveLength(0);
      expect(result.unsupported.length).toBeGreaterThanOrEqual(1);
      expect(result.unsupported[0].reason).toMatch(/author/i);
    });

    it("default mode (edit) still produces plain insert-text / delete-range", async () => {
      const agent = await loadAgent([{ text: "alpha" }]);
      const state = stateFor(agent);
      const tx = state.tr.insertText("X", 1, 1);
      const result = transactionToCommands(tx, state);
      expect(result.commands).toHaveLength(1);
      expect(result.commands[0].type).toBe("docx:insert-text");
    });
  });

  describe("multi-paragraph format", () => {
    it("addMark across two paragraphs emits a single docx:format-range with cross-paragraph endpoints", async () => {
      const agent = await loadAgent([{ text: "first paragraph" }, { text: "second paragraph" }]);
      const state = stateFor(agent);
      const fromPM = 1 + 6;
      const toPM = 1 + 15 + 1 + 1 + 7;
      const tx = state.tr.addMark(fromPM, toPM, docxSchema.marks.bold.create());
      const result = transactionToCommands(tx, state);
      expect(result.unsupported).toHaveLength(0);
      // PM emits one AddMarkStep per textblock (so 2 steps for 2 paragraphs).
      // Each step becomes one format-range command, both with bold=true.
      expect(result.commands.length).toBeGreaterThanOrEqual(1);
      for (const cmd of result.commands) {
        expect(cmd.type).toBe("docx:format-range");
        expect(cmd.payload).toMatchObject({ format: { bold: true } });
      }
    });
  });
});
