import { describe, expect, it } from "vitest";
import { EditorState, TextSelection } from "prosemirror-state";
import { DocxAgent, docToPM, docxSchema } from "@officeai/docx";
import { buildSampleDocx } from "./sample-docx";
import { dispatchShortcut } from "./word-shortcuts-keymap";

/**
 * Unit tests for the Word-parity shortcut keymap.
 *
 * Covers the regression where `Mod-B` on already-bold text would
 * re-apply `bold: true` instead of toggling it off (the snapshot
 * probe used `(paragraph, offset)` and mis-read the boundary case
 * where the selection start sat between an unmarked and a marked
 * run). The fix moved the toggle decision over to PM's
 * `activeMarks` set; these tests pin that contract.
 */

async function loadAgent(): Promise<DocxAgent> {
  const buf = await buildSampleDocx();
  return DocxAgent.fromBuffer(buf);
}

/**
 * Find the (paragraphIndex, runIndex) that contains a given substring
 * in its text leaf. We use this to make assertions independent of the
 * sample document's internal ordering — the only thing the keymap
 * needs to be right about is "the run we asked for".
 */
function locateRunByText(agent: DocxAgent, needle: string): { paragraphIndex: number; runIndex: number } {
  const body = agent.getSnapshot().root.body;
  for (let p = 0; p < body.length; p++) {
    const block = body[p];
    if (block?.kind !== "paragraph") continue;
    for (let r = 0; r < block.children.length; r++) {
      const child = block.children[r];
      if (child?.kind !== "run") continue;
      const txt = child.children.map((leaf) => (leaf.kind === "text" ? leaf.text : "")).join("");
      if (txt.includes(needle)) return { paragraphIndex: p, runIndex: r };
    }
  }
  throw new Error(`no run containing "${needle}"`);
}

/** Compute the PM from/to range for the entire first paragraph's text. */
function selectFirstParagraph(state: EditorState): { from: number; to: number } {
  let from = -1;
  let to = -1;
  state.doc.descendants((node, pos) => {
    if (from >= 0) return false;
    if (node.type.name === "paragraph") {
      from = pos + 1;
      to = pos + node.nodeSize - 1;
      return false;
    }
    return true;
  });
  if (from < 0 || to <= from) throw new Error("no paragraph in PM doc");
  return { from, to };
}

function buildView(agent: DocxAgent): {
  state: EditorState;
  dom: { dispatchEvent: () => boolean };
} {
  const state = EditorState.create({
    schema: docxSchema,
    doc: docToPM(agent.getSnapshot()),
  });
  return { state, dom: { dispatchEvent: () => true } };
}

function modKeyEvent(key: string, modifiers: { shift?: boolean; alt?: boolean } = {}): KeyboardEvent {
  return {
    key,
    code: key.startsWith("Digit") ? key : key === "Enter" ? "Enter" : `Key${key.toUpperCase()}`,
    metaKey: true,
    ctrlKey: false,
    shiftKey: modifiers.shift ?? false,
    altKey: modifiers.alt ?? false,
    preventDefault: () => undefined,
    stopPropagation: () => undefined,
  } as unknown as KeyboardEvent;
}

async function flushBus(): Promise<void> {
  // The keymap fires `void agent.applyCommand(...)` so the mutation
  // resolves on the microtask queue. Awaiting twice is enough to drain
  // both the dispatch promise and the subscribe callback chain.
  await Promise.resolve();
  await Promise.resolve();
}

describe("word-shortcuts-keymap — dispatchShortcut", () => {
  describe("Mod-B toggle", () => {
    it("applies bold when no run in the selection carries it", async () => {
      const agent = await loadAgent();
      const view = buildView(agent);
      const range = selectFirstParagraph(view.state);
      view.state = view.state.apply(
        view.state.tr.setSelection(TextSelection.create(view.state.doc, range.from, range.to))
      );

      const handled = dispatchShortcut(view as never, modKeyEvent("b"), agent);
      await flushBus();

      expect(handled).toBe(true);
      const block = agent.getSnapshot().root.body[0];
      if (block?.kind !== "paragraph") throw new Error("expected paragraph");
      const someRunBold = block.children.some((c) => c.kind === "run" && c.properties.bold === true);
      expect(someRunBold).toBe(true);
    });

    it("removes bold when the entire selection is already bold (regression: cmd+b twice should toggle off)", async () => {
      const agent = await loadAgent();
      // The sample doc's first paragraph is the title. Make the whole
      // first paragraph explicitly bold via the agent so PM projects
      // `<strong>` over every run.
      const block0 = agent.getSnapshot().root.body[0];
      if (block0?.kind !== "paragraph") throw new Error("expected paragraph");
      const flatLen = block0.children.reduce((n, c) => {
        if (c.kind !== "run") return n;
        return n + c.children.reduce((m, leaf) => m + (leaf.kind === "text" ? leaf.text.length : 0), 0);
      }, 0);
      await agent.applyCommand({
        type: "docx:format-range",
        payload: {
          range: {
            start: { paragraph: 0, run: 0, offset: 0 },
            end: { paragraph: 0, run: 0, offset: flatLen },
          },
          format: { bold: true },
        },
        source: "human",
      });

      const view = buildView(agent);
      const range = selectFirstParagraph(view.state);
      view.state = view.state.apply(
        view.state.tr.setSelection(TextSelection.create(view.state.doc, range.from, range.to))
      );

      const handled = dispatchShortcut(view as never, modKeyEvent("b"), agent);
      await flushBus();

      expect(handled).toBe(true);
      const blockAfter = agent.getSnapshot().root.body[0];
      if (blockAfter?.kind !== "paragraph") throw new Error("expected paragraph");
      const anyRunBold = blockAfter.children.some((c) => c.kind === "run" && c.properties.bold === true);
      // The fix: bold must be cleared on every run, not re-applied.
      expect(anyRunBold).toBe(false);
    });

    it("returns false (no-op) when the selection is collapsed", async () => {
      const agent = await loadAgent();
      const view = buildView(agent);
      view.state = view.state.apply(view.state.tr.setSelection(TextSelection.create(view.state.doc, 2, 2)));
      const handled = dispatchShortcut(view as never, modKeyEvent("b"), agent);
      expect(handled).toBe(false);
    });
  });

  describe("alignment chords — Google-Docs parity (Mod+Shift+L/E/R/J)", () => {
    // Bare-modifier chords MUST fall through so the browser's own
    // shortcuts win — Cmd+R reload, Cmd+L address bar, etc. Pin
    // every alignment letter so a regression here surfaces in CI
    // instead of as a bug report ("the page won't refresh!").
    it.each([["l"], ["e"], ["r"], ["j"]])(
      "Mod+%s alone returns false (browser keeps the shortcut)",
      async (k) => {
        const agent = await loadAgent();
        const view = buildView(agent);
        const range = selectFirstParagraph(view.state);
        view.state = view.state.apply(
          view.state.tr.setSelection(TextSelection.create(view.state.doc, range.from, range.to))
        );
        const handled = dispatchShortcut(view as never, modKeyEvent(k), agent);
        expect(handled).toBe(false);
      }
    );

    it.each([
      ["l", "left"],
      ["e", "center"],
      ["r", "right"],
      ["j", "justify"],
    ] as const)("Mod+Shift+%s applies %s alignment", async (k, alignment) => {
      const agent = await loadAgent();
      const view = buildView(agent);
      const range = selectFirstParagraph(view.state);
      view.state = view.state.apply(
        view.state.tr.setSelection(TextSelection.create(view.state.doc, range.from, range.to))
      );
      const handled = dispatchShortcut(view as never, modKeyEvent(k, { shift: true }), agent);
      await flushBus();
      expect(handled).toBe(true);
      const block = agent.getSnapshot().root.body[0];
      if (block?.kind !== "paragraph") throw new Error("expected paragraph");
      expect(block.properties.alignment).toBe(alignment);
    });
  });

  describe("Mod-Shift-X toggle (strikethrough)", () => {
    it("maps the PM mark name `strikethrough` onto the OOXML `strike` field", async () => {
      const agent = await loadAgent();
      // Pre-apply strike across the first paragraph so PM projects an
      // <s> span; pressing the shortcut should clear the `strike` flag.
      const block0 = agent.getSnapshot().root.body[0];
      if (block0?.kind !== "paragraph") throw new Error("expected paragraph");
      const flatLen = block0.children.reduce((n, c) => {
        if (c.kind !== "run") return n;
        return n + c.children.reduce((m, leaf) => m + (leaf.kind === "text" ? leaf.text.length : 0), 0);
      }, 0);
      await agent.applyCommand({
        type: "docx:format-range",
        payload: {
          range: {
            start: { paragraph: 0, run: 0, offset: 0 },
            end: { paragraph: 0, run: 0, offset: flatLen },
          },
          format: { strike: true },
        },
        source: "human",
      });

      const view = buildView(agent);
      const range = selectFirstParagraph(view.state);
      view.state = view.state.apply(
        view.state.tr.setSelection(TextSelection.create(view.state.doc, range.from, range.to))
      );

      const handled = dispatchShortcut(view as never, modKeyEvent("x", { shift: true }), agent);
      await flushBus();

      expect(handled).toBe(true);
      const blockAfter = agent.getSnapshot().root.body[0];
      if (blockAfter?.kind !== "paragraph") throw new Error("expected paragraph");
      const anyStrike = blockAfter.children.some((c) => c.kind === "run" && c.properties.strike === true);
      expect(anyStrike).toBe(false);
    });
  });
});

void locateRunByText;
