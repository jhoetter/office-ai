/**
 * @vitest-environment jsdom
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  handleUndoRedo,
  isFormField,
  isRedoChord,
  isUndoChord,
  runRedo,
  runUndo,
  type UndoableAgent,
} from "./undo-redo";

/**
 * The shared undo / redo helper is the seam between the renderer
 * keymaps and the headless `CommandBus`. These tests pin its
 * behaviour so a regression here would immediately surface as
 * "Cmd+Z stopped working in some editor" — the exact symptom the
 * unify-undo-redo-stacks plan is trying to make impossible.
 */

function makeKey(
  key: string,
  mods: { meta?: boolean; ctrl?: boolean; shift?: boolean; alt?: boolean } = {},
  target?: EventTarget | null
): KeyboardEvent {
  const event = new KeyboardEvent("keydown", {
    key,
    metaKey: mods.meta ?? false,
    ctrlKey: mods.ctrl ?? false,
    shiftKey: mods.shift ?? false,
    altKey: mods.alt ?? false,
    bubbles: true,
    cancelable: true,
  });
  if (target) {
    Object.defineProperty(event, "target", { value: target });
  }
  return event;
}

function makeAgent(opts: { canUndo?: boolean; canRedo?: boolean } = {}): UndoableAgent & {
  undo: ReturnType<typeof vi.fn>;
  redo: ReturnType<typeof vi.fn>;
} {
  return {
    canUndo: () => opts.canUndo ?? true,
    canRedo: () => opts.canRedo ?? true,
    undo: vi.fn(),
    redo: vi.fn(),
  };
}

describe("isUndoChord", () => {
  it.each([
    ["Cmd+Z", { meta: true }],
    ["Ctrl+Z", { ctrl: true }],
    ["Cmd+z (lowercase key)", { meta: true }],
  ])("matches %s", (_label, mods) => {
    expect(isUndoChord(makeKey("z", mods))).toBe(true);
    expect(isUndoChord(makeKey("Z", mods))).toBe(true);
  });

  it("rejects Cmd+Shift+Z (that's redo)", () => {
    expect(isUndoChord(makeKey("z", { meta: true, shift: true }))).toBe(false);
  });

  it("rejects modifier-less Z", () => {
    expect(isUndoChord(makeKey("z"))).toBe(false);
  });

  it("rejects Cmd+Alt+Z to avoid stomping system shortcuts", () => {
    expect(isUndoChord(makeKey("z", { meta: true, alt: true }))).toBe(false);
  });
});

describe("isRedoChord", () => {
  it.each([
    ["Cmd+Shift+Z", { meta: true, shift: true }],
    ["Ctrl+Shift+Z", { ctrl: true, shift: true }],
  ])("matches %s", (_label, mods) => {
    expect(isRedoChord(makeKey("z", mods))).toBe(true);
  });

  it.each([
    ["Cmd+Y", { meta: true }],
    ["Ctrl+Y", { ctrl: true }],
  ])("matches %s", (_label, mods) => {
    expect(isRedoChord(makeKey("y", mods))).toBe(true);
    expect(isRedoChord(makeKey("Y", mods))).toBe(true);
  });

  it("rejects Cmd+Z (that's undo)", () => {
    expect(isRedoChord(makeKey("z", { meta: true }))).toBe(false);
  });

  it("rejects Cmd+Shift+Y to avoid colliding with redo-via-Y users' muscle memory", () => {
    // Cmd+Shift+Y isn't a standard chord; if we accepted it we'd
    // start redoing in surfaces that bind Cmd+Shift+Y to their own
    // thing (e.g. some PDF viewers' "show search").
    expect(isRedoChord(makeKey("y", { meta: true, shift: true }))).toBe(false);
  });
});

describe("isFormField", () => {
  let host: HTMLDivElement;

  beforeEach(() => {
    host = document.createElement("div");
    document.body.appendChild(host);
  });

  afterEach(() => {
    host.remove();
  });

  it("blocks undo on <input>", () => {
    const input = document.createElement("input");
    host.appendChild(input);
    expect(isFormField(input)).toBe(true);
  });

  it("blocks undo on <textarea>", () => {
    const ta = document.createElement("textarea");
    host.appendChild(ta);
    expect(isFormField(ta)).toBe(true);
  });

  it("blocks undo on a contenteditable element outside the editor surface", () => {
    const ce = document.createElement("div");
    ce.setAttribute("contenteditable", "true");
    host.appendChild(ce);
    expect(isFormField(ce)).toBe(true);
  });

  it("does NOT block undo when the contenteditable is the editor surface itself", () => {
    const editor = document.createElement("div");
    editor.setAttribute("contenteditable", "true");
    host.appendChild(editor);
    expect(isFormField(editor, editor)).toBe(false);
  });

  it("does NOT block undo when the target is descended from the editor surface", () => {
    const editor = document.createElement("div");
    editor.setAttribute("contenteditable", "true");
    const inner = document.createElement("p");
    inner.setAttribute("contenteditable", "true");
    editor.appendChild(inner);
    host.appendChild(editor);
    expect(isFormField(inner, editor)).toBe(false);
  });

  it("ignores non-element targets gracefully", () => {
    expect(isFormField(null)).toBe(false);
    expect(isFormField(window)).toBe(false);
  });
});

describe("runUndo / runRedo", () => {
  it("runUndo calls agent.undo when canUndo is true", () => {
    const agent = makeAgent({ canUndo: true });
    expect(runUndo(agent)).toBe(true);
    expect(agent.undo).toHaveBeenCalledTimes(1);
  });

  it("runUndo is a no-op when canUndo is false", () => {
    const agent = makeAgent({ canUndo: false });
    expect(runUndo(agent)).toBe(false);
    expect(agent.undo).not.toHaveBeenCalled();
  });

  it("runUndo tolerates a null agent (renders haven't mounted yet)", () => {
    expect(runUndo(null)).toBe(false);
    expect(runUndo(undefined)).toBe(false);
  });

  it("runRedo mirrors runUndo for redo", () => {
    const agent = makeAgent({ canRedo: true });
    expect(runRedo(agent)).toBe(true);
    expect(agent.redo).toHaveBeenCalledTimes(1);

    const stuck = makeAgent({ canRedo: false });
    expect(runRedo(stuck)).toBe(false);
    expect(stuck.redo).not.toHaveBeenCalled();
  });
});

describe("handleUndoRedo", () => {
  it("dispatches undo, calls preventDefault, returns true", () => {
    const agent = makeAgent();
    const e = makeKey("z", { meta: true });
    const prevent = vi.spyOn(e, "preventDefault");
    expect(handleUndoRedo(e, agent)).toBe(true);
    expect(prevent).toHaveBeenCalled();
    expect(agent.undo).toHaveBeenCalledTimes(1);
    expect(agent.redo).not.toHaveBeenCalled();
  });

  it("dispatches redo on Cmd+Shift+Z", () => {
    const agent = makeAgent();
    const e = makeKey("z", { meta: true, shift: true });
    expect(handleUndoRedo(e, agent)).toBe(true);
    expect(agent.redo).toHaveBeenCalledTimes(1);
    expect(agent.undo).not.toHaveBeenCalled();
  });

  it("dispatches redo on Cmd+Y", () => {
    const agent = makeAgent();
    const e = makeKey("y", { meta: true });
    expect(handleUndoRedo(e, agent)).toBe(true);
    expect(agent.redo).toHaveBeenCalledTimes(1);
  });

  it("returns false for unrelated chords without touching the agent", () => {
    const agent = makeAgent();
    const e = makeKey("a", { meta: true });
    const prevent = vi.spyOn(e, "preventDefault");
    expect(handleUndoRedo(e, agent)).toBe(false);
    expect(prevent).not.toHaveBeenCalled();
    expect(agent.undo).not.toHaveBeenCalled();
    expect(agent.redo).not.toHaveBeenCalled();
  });

  it("does NOT preventDefault when the focus is in a form field — browser undo wins", () => {
    const input = document.createElement("input");
    document.body.appendChild(input);
    try {
      const agent = makeAgent();
      const e = makeKey("z", { meta: true }, input);
      const prevent = vi.spyOn(e, "preventDefault");
      expect(handleUndoRedo(e, agent)).toBe(false);
      expect(prevent).not.toHaveBeenCalled();
      expect(agent.undo).not.toHaveBeenCalled();
    } finally {
      input.remove();
    }
  });

  it("DOES dispatch undo when focus is inside the editor surface (even though it's contenteditable)", () => {
    const editor = document.createElement("div");
    editor.setAttribute("contenteditable", "true");
    document.body.appendChild(editor);
    try {
      const agent = makeAgent();
      const e = makeKey("z", { meta: true }, editor);
      expect(handleUndoRedo(e, agent, { editorSurface: editor })).toBe(true);
      expect(agent.undo).toHaveBeenCalledTimes(1);
    } finally {
      editor.remove();
    }
  });

  it("returns true (preventDefault'd) but does not dispatch when canUndo is false", () => {
    // Important: we still preventDefault so the browser doesn't pop
    // up its own native undo on the editor surface — the user's
    // intent was clearly "undo the editor", not "undo a stray
    // browser-level edit". This matches what the inline handlers
    // we're replacing always did.
    const agent = makeAgent({ canUndo: false });
    const e = makeKey("z", { meta: true });
    const prevent = vi.spyOn(e, "preventDefault");
    expect(handleUndoRedo(e, agent)).toBe(false);
    expect(prevent).toHaveBeenCalled();
    expect(agent.undo).not.toHaveBeenCalled();
  });
});
