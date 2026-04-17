import { EditorState, type Transaction } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { keymap } from "prosemirror-keymap";
import { baseKeymap } from "prosemirror-commands";
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
 * Mount a ProseMirror EditorView whose `dispatchTransaction` funnels every
 * change into `agent.applyCommand(s)` instead of mutating the view directly.
 *
 * The view receives DOM-side updates through an `agent.subscribe` callback
 * that re-projects the snapshot via `docToPM` and sets a `from-bus` meta
 * marker so the funnel knows to pass the transaction through.
 *
 * Browser-only: requires a DOM and an `EditorView` host element.
 */
export function mountDocxEditor(target: Element, opts: MountOptions): MountResult {
  const { agent } = opts;
  const initialDoc = docToPM(agent.getSnapshot());
  const state = EditorState.create({
    schema: docxSchema,
    doc: initialDoc,
    plugins: [keymap(baseKeymap)],
  });

  let isProjecting = false;

  const view = new EditorView(target, {
    state,
    dispatchTransaction(tx: Transaction) {
      if (tx.getMeta("from-bus") === true || isProjecting) {
        view.updateState(view.state.apply(tx));
        return;
      }
      if (tx.steps.length === 0) {
        view.updateState(view.state.apply(tx));
        return;
      }
      const result = transactionToCommands(tx, view.state, {
        ...(opts.agentId !== undefined ? { agentId: opts.agentId } : {}),
        source: opts.source ?? "human",
      });
      if (result.unsupported.length > 0) opts.onUnsupported?.(result.unsupported);
      if (result.commands.length === 0) return;
      void agent.applyCommands(result.commands).catch((err) => opts.onError?.(err));
    },
  });

  const unsubscribe = agent.subscribe(() => {
    isProjecting = true;
    try {
      const pm = docToPM(agent.getSnapshot());
      const tx = view.state.tr.replaceWith(0, view.state.doc.content.size, pm.content);
      tx.setMeta("from-bus", true);
      view.dispatch(tx);
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
