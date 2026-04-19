import { describe, expect, it, vi } from "vitest";
import { EditorState, Plugin, TextSelection } from "prosemirror-state";
import { DocxAgent, docToPM, docxSchema } from "@officeai/docx";
import { buildSampleDocx } from "./sample-docx";
import { pageKeymapPlugin } from "./page-keymap";

/**
 * Unit tests for the page-aware keymap.
 *
 * The most important contract: pressing Mod+Enter inside a body
 * paragraph dispatches `docx:insert-page-break` against the agent.
 * The chunker downstream then re-paginates on the next snapshot —
 * that's covered by the e2e suite, here we just pin the handler
 * → command wire-up so a regression in the keymap can't slip
 * through silently (the previous bug was masked because the only
 * coverage was an e2e test that needed Playwright + a focused
 * editor surface, both of which are easy to break in CI).
 */

async function loadAgent(): Promise<DocxAgent> {
  const buf = await buildSampleDocx();
  return DocxAgent.fromBuffer(buf);
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

function placeCaretInFirstParagraph(state: EditorState): EditorState {
  let pos = -1;
  state.doc.descendants((node, p) => {
    if (pos >= 0) return false;
    if (node.type.name === "paragraph") {
      pos = p + 1;
      return false;
    }
    return true;
  });
  if (pos < 0) throw new Error("no paragraph in PM doc");
  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));
}

function modEnter(): KeyboardEvent {
  return {
    key: "Enter",
    code: "Enter",
    metaKey: true,
    ctrlKey: false,
    shiftKey: false,
    altKey: false,
    preventDefault: () => undefined,
    stopPropagation: () => undefined,
  } as unknown as KeyboardEvent;
}

/** Pull the keydown handler off a built plugin so we can call it without a real EditorView. */
function keydownOf(plugin: Plugin): (view: unknown, event: KeyboardEvent) => boolean {
  // The keymap plugin registers handleKeyDown via PM's Plugin spec;
  // PM exposes plugin props through `.props`. We grab it directly so
  // the test stays decoupled from the plugin's internal layout.
  const props = (plugin as unknown as { props: { handleKeyDown?: typeof handler } }).props;
  type Handler = (view: unknown, event: KeyboardEvent) => boolean;
  const handler = props.handleKeyDown as Handler | undefined;
  if (!handler) throw new Error("plugin has no handleKeyDown");
  return handler;
}

describe("page-keymap — handleKeyDown", () => {
  it("Mod+Enter dispatches docx:insert-page-break against the caret paragraph", async () => {
    const agent = await loadAgent();
    const view = buildView(agent);
    view.state = placeCaretInFirstParagraph(view.state);

    const spy = vi.spyOn(agent, "applyCommand");
    const plugin = pageKeymapPlugin(agent);
    const handled = keydownOf(plugin)(view, modEnter());

    expect(handled).toBe(true);
    expect(spy).toHaveBeenCalledTimes(1);
    const arg = spy.mock.calls[0]?.[0] as { type: string; payload: { paragraphId: string; offset: number } };
    expect(arg.type).toBe("docx:insert-page-break");
    expect(typeof arg.payload.paragraphId).toBe("string");
    expect(arg.payload.paragraphId.length).toBeGreaterThan(0);
    expect(arg.payload.offset).toBe(0);
  });

  it("Mod+Enter actually mutates the snapshot — a page-break leaf lands in the body", async () => {
    const agent = await loadAgent();
    const view = buildView(agent);
    view.state = placeCaretInFirstParagraph(view.state);

    const before = countPageBreaks(agent);
    const plugin = pageKeymapPlugin(agent);
    const handled = keydownOf(plugin)(view, modEnter());
    expect(handled).toBe(true);
    // applyCommand resolves on the microtask queue; drain it.
    await Promise.resolve();
    await Promise.resolve();

    const after = countPageBreaks(agent);
    expect(after).toBe(before + 1);
  });
});

/**
 * Count `pageBreak` leaves across all body paragraphs. Used to assert the
 * keymap genuinely produces a structural mutation rather than just
 * dispatching a command that the handler then drops on the floor.
 */
function countPageBreaks(agent: DocxAgent): number {
  const body = agent.getSnapshot().root.body;
  let n = 0;
  for (const block of body) {
    if (block.kind !== "paragraph") continue;
    for (const child of block.children) {
      if (child.kind !== "run") continue;
      for (const leaf of child.children) {
        if (leaf.kind === "page-break") n++;
      }
    }
  }
  return n;
}
