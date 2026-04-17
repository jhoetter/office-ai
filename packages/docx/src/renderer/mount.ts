import { EditorState, TextSelection, type Transaction } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { keymap } from "prosemirror-keymap";
import { baseKeymap } from "prosemirror-commands";
import { history, undo, redo } from "prosemirror-history";
import type { DocxAgent } from "../agent/agent.js";
import { docxSchema } from "./schema.js";
import { docToPM } from "./doc-to-pm.js";
import { transactionToCommands, type UnsupportedTx } from "./transaction-to-commands.js";

export interface MountOptions {
  agent: DocxAgent;
  agentId?: string;
  source?: "human" | "agent" | "system";
  onUnsupported?: (events: UnsupportedTx[]) => void;
  onError?: (err: unknown) => void;
}

export interface MountResult {
  view: EditorView;
  destroy: () => void;
}

/**
 * Mount a ProseMirror EditorView whose `dispatchTransaction` mirrors every
 * change into `agent.applyCommand(s)` (the headless command bus).
 *
 * Architecture:
 *
 *   1. The user's transaction is applied to the EditorView **immediately**
 *      so the DOM and selection update synchronously and typing feels
 *      native. The bus is mirrored in the background (fire-and-forget).
 *
 *   2. When a mutation arrives at the bus that did *not* originate from
 *      this view (e.g. an agent prompt, a comment dispatched outside the
 *      funnel, an external collaborator), the snapshot is re-projected
 *      into PM and the previous selection is mapped through the change.
 *
 * This design keeps the editor responsive while preserving the invariant
 * "every mutation flows through the bus".
 */
export function mountDocxEditor(target: Element, opts: MountOptions): MountResult {
  const { agent } = opts;
  const initialDoc = docToPM(agent.getSnapshot());
  const state = EditorState.create({
    schema: docxSchema,
    doc: initialDoc,
    plugins: [
      history(),
      keymap({
        "Mod-z": undo,
        "Mod-y": redo,
        "Mod-Shift-z": redo,
      }),
      keymap(baseKeymap),
    ],
  });

  // Counts mutations that the funnel has dispatched into the bus and is
  // waiting to "echo" through the subscribe callback. We use a counter
  // (not a boolean) because handlers run synchronously inside
  // `applyCommands`, so subscribe fires once per command.
  let pendingFunnelCount = 0;
  let isProjecting = false;

  const view = new EditorView(target, {
    state,
    dispatchTransaction(tx: Transaction) {
      // Projections initiated by the bus go through unmodified.
      if (tx.getMeta("from-bus") === true || isProjecting) {
        view.updateState(view.state.apply(tx));
        return;
      }

      const before = view.state;
      // Apply the user's transaction immediately. This is what keeps the
      // cursor where the user expects it and what makes typing feel native.
      view.updateState(before.apply(tx));

      if (tx.steps.length === 0) return;

      const result = transactionToCommands(tx, before, {
        ...(opts.agentId !== undefined ? { agentId: opts.agentId } : {}),
        source: opts.source ?? "human",
      });
      if (result.unsupported.length > 0) opts.onUnsupported?.(result.unsupported);
      if (result.commands.length === 0) return;

      // Mark these as "self-originated" so the subscribe callback below
      // does not echo them back as a re-projection (which would clobber
      // the selection we just preserved).
      pendingFunnelCount += result.commands.length;
      void agent.applyCommands(result.commands).catch((err) => {
        // Make sure we don't leak the suppression count if the bus
        // rejected our commands. Subscribe will not have fired for
        // rejected commands (handler errors emit an internal mutation
        // but still call notify), so treat this as a hard reset.
        pendingFunnelCount = Math.max(0, pendingFunnelCount - result.commands.length);
        opts.onError?.(err);
      });
    },
  });

  const unsubscribe = agent.subscribe(() => {
    if (pendingFunnelCount > 0) {
      pendingFunnelCount--;
      return;
    }

    // External mutation (e.g. agent prompt, direct agent.applyCommand from
    // the host UI). Re-project the snapshot and map the selection through
    // the structural change so the cursor doesn't jump to position 0.
    isProjecting = true;
    try {
      const oldSelectionAnchor = view.state.selection.from;
      const oldSelectionHead = view.state.selection.to;
      const pm = docToPM(agent.getSnapshot());
      const tr = view.state.tr.replaceWith(0, view.state.doc.content.size, pm.content);
      // Best-effort: clamp the previous selection into the new doc.
      const docSize = tr.doc.content.size;
      const anchor = clampSelection(tr.doc, Math.min(oldSelectionAnchor, docSize));
      const head = clampSelection(tr.doc, Math.min(oldSelectionHead, docSize));
      try {
        tr.setSelection(TextSelection.create(tr.doc, anchor, head));
      } catch {
        /* selection couldn't be set; PM will fall back to default */
      }
      tr.setMeta("from-bus", true);
      view.dispatch(tr);
    } finally {
      isProjecting = false;
    }
  });

  return {
    view,
    destroy() {
      unsubscribe();
      view.destroy();
    },
  };
}

function clampSelection(doc: import("prosemirror-model").Node, pos: number): number {
  const max = doc.content.size;
  let p = Math.max(0, Math.min(max, pos));
  // Walk to a valid text-cursor position (PM rejects positions that
  // aren't inside a textblock).
  const $pos = doc.resolve(p);
  if (!$pos.parent.isTextblock) {
    // Find nearest textblock by scanning forward then backward.
    for (let i = p; i <= max; i++) {
      if (doc.resolve(i).parent.isTextblock) return i;
    }
    for (let i = p; i >= 0; i--) {
      if (doc.resolve(i).parent.isTextblock) return i;
    }
  }
  return p;
}
