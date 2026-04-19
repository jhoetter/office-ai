// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { Slice, Fragment } from "prosemirror-model";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { makeSyntheticDocx, plainDocxXml } from "../test-utils/synthetic.js";
import { mountDocxEditor } from "./mount.js";
import { docxSchema } from "./schema.js";
import type { Paragraph } from "../model/types.js";

/**
 * Mount-level tests for the DOCX renderer. These pin the
 * "bus is the only undo source of truth" invariants set out in
 * spec/shared/agent-api.md so they survive future refactors.
 *
 * Why a separate suite from `transaction-to-commands.test.ts`?
 *   - Those tests exercise translation in isolation.
 *   - These exercise the real `EditorView` + bus wiring, where
 *     keymap routing, projection metadata, and drift guards
 *     actually matter.
 */

async function loadAgent(): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({
    documentXml: plainDocxXml([{ text: "hello world" }, { text: "second paragraph" }]),
  });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

let host: HTMLDivElement;

beforeEach(() => {
  host = document.createElement("div");
  document.body.appendChild(host);
});

afterEach(() => {
  host.remove();
});

/**
 * Synthesise the keydown event PM's keymap plugin actually
 * inspects. `prosemirror-keymap` normalises the `"Mod-"` prefix
 * to either `Meta` (on Mac) or `Ctrl` (everywhere else) based on
 * `navigator.platform`, then looks up the binding by EXACTLY
 * matching the modifier string ("Meta-z", "Ctrl-z", etc). Setting
 * both modifiers produces "Meta-Ctrl-z" which matches neither —
 * so we mirror PM's own platform check and set only the right
 * one.
 */
const isMacPlatform =
  typeof navigator !== "undefined" && /Mac|iP(hone|[oa]d)/.test(navigator.platform);
function modKey(key: string, opts: { shift?: boolean } = {}): KeyboardEvent {
  return new KeyboardEvent("keydown", {
    key,
    metaKey: isMacPlatform,
    ctrlKey: !isMacPlatform,
    shiftKey: opts.shift ?? false,
    bubbles: true,
    cancelable: true,
  });
}

describe("mountDocxEditor — undo / redo keymap routes through the bus", () => {
  it("Mod-Z calls agent.undo when there's something to undo", async () => {
    const agent = await loadAgent();
    const mount = mountDocxEditor(host, { agent });
    try {
      // Apply something the bus can undo (typing through the bus
      // creates an "approved" mutation). Going through the agent
      // keeps this test independent of dispatchTransaction's
      // optimistic-apply path.
      const para = agent.getSnapshot().root.body[0] as Paragraph;
      await agent.applyCommand({
        type: "docx:insert-text",
        payload: { at: { paragraph: 0, run: 0, offset: para ? 0 : 0 }, text: "X" },
        source: "human",
      });
      expect(agent.canUndo()).toBe(true);

      const undo = vi.spyOn(agent, "undo");
      const handled = mount.view.someProp("handleKeyDown", (fn) =>
        fn(mount.view, modKey("z"))
      );
      expect(handled).toBe(true);
      expect(undo).toHaveBeenCalledTimes(1);
    } finally {
      mount.destroy();
    }
  });

  it("Mod-Shift-Z calls agent.redo when canRedo is true", async () => {
    const agent = await loadAgent();
    const mount = mountDocxEditor(host, { agent });
    try {
      await agent.applyCommand({
        type: "docx:insert-text",
        payload: { at: { paragraph: 0, run: 0, offset: 0 }, text: "X" },
        source: "human",
      });
      agent.undo();
      expect(agent.canRedo()).toBe(true);

      const redo = vi.spyOn(agent, "redo");
      const handled = mount.view.someProp("handleKeyDown", (fn) =>
        fn(mount.view, modKey("z", { shift: true }))
      );
      expect(handled).toBe(true);
      expect(redo).toHaveBeenCalledTimes(1);
    } finally {
      mount.destroy();
    }
  });

  it("Mod-Z still consumes the keystroke when canUndo is false (no browser-level undo leak)", async () => {
    const agent = await loadAgent();
    const mount = mountDocxEditor(host, { agent });
    try {
      expect(agent.canUndo()).toBe(false);
      const undo = vi.spyOn(agent, "undo");
      const handled = mount.view.someProp("handleKeyDown", (fn) =>
        fn(mount.view, modKey("z"))
      );
      // The PM keymap consumes the event (returns true) so the
      // browser doesn't fire its own contenteditable undo on the
      // editor surface — that was a known phantom-edit source.
      expect(handled).toBe(true);
      // ...but the bus didn't get called because there's nothing
      // to undo.
      expect(undo).not.toHaveBeenCalled();
    } finally {
      mount.destroy();
    }
  });
});

describe("mountDocxEditor — from-bus projection metadata", () => {
  it("an external agent.applyCommand projects with from-bus + addToHistory:false on the dispatched tx", async () => {
    const agent = await loadAgent();
    const mount = mountDocxEditor(host, { agent });
    try {
      // Capture the very next dispatched transaction by patching
      // dispatch. We can't `view.dispatch = ...` directly because
      // the bus subscribe handler reads it via closure on the
      // EditorView's `dispatchTransaction` slot, but we CAN
      // intercept via prosemirror's own override slot.
      const dispatched: Array<{ fromBus: unknown; addToHistory: unknown }> = [];
      const originalDispatch = mount.view.dispatch.bind(mount.view);
      mount.view.dispatch = (tr) => {
        dispatched.push({
          fromBus: tr.getMeta("from-bus"),
          addToHistory: tr.getMeta("addToHistory"),
        });
        originalDispatch(tr);
      };

      // Drive a snapshot mutation that did NOT come through the
      // mount's funnel — exactly the case where the subscribe
      // handler re-projects the snapshot.
      await agent.applyCommand({
        type: "docx:insert-text",
        payload: { at: { paragraph: 0, run: 0, offset: 0 }, text: "AGENT " },
        source: "agent",
        agentId: "test",
      });

      const projection = dispatched.find((d) => d.fromBus === true);
      expect(projection, "expected a from-bus projection transaction").toBeDefined();
      // The invariant: the bus owns the canonical history, so the
      // projection must NEVER end up on a renderer history stack.
      expect(projection?.addToHistory).toBe(false);
    } finally {
      mount.destroy();
    }
  });
});

describe("mountDocxEditor — drift guard: unsupported transactions don't mutate the PM doc", () => {
  it("a transaction whose steps map to zero commands leaves the PM doc unchanged and pings onUnsupported", async () => {
    const agent = await loadAgent();
    const onUnsupported = vi.fn();
    const mount = mountDocxEditor(host, { agent, onUnsupported });
    try {
      const beforeDoc = mount.view.state.doc;

      // Construct a transaction whose only step inserts a `table`
      // atom into the body. `transaction-to-commands.ts` flags any
      // slice containing a non-paragraph block as
      // `slice contains a non-paragraph block (table / opaque)` —
      // it produces zero commands AND an `unsupported` event.
      // Pre-fix this path silently mutated the PM doc, leaving the
      // bus snapshot stale; the new drift guard refuses to apply.
      const tableNode = docxSchema.nodes.table.create({});
      const slice = new Slice(Fragment.from(tableNode), 0, 0);
      const tr = mount.view.state.tr.replace(0, 0, slice);
      mount.view.dispatch(tr);

      // The doc must equal the pre-dispatch doc — no drift.
      expect(mount.view.state.doc.eq(beforeDoc)).toBe(true);
      expect(onUnsupported).toHaveBeenCalledTimes(1);
    } finally {
      mount.destroy();
    }
  });
});
