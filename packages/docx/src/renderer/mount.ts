import { EditorState, Plugin, TextSelection, type Transaction } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { keymap } from "prosemirror-keymap";
import { baseKeymap } from "prosemirror-commands";
import type { DocxAgent } from "../agent/agent.js";
import { docxSchema } from "./schema.js";
import { docToPM } from "./doc-to-pm.js";
import { transactionToCommands, type UnsupportedTx } from "./transaction-to-commands.js";

/**
 * Why no `prosemirror-history` plugin?
 *
 * Every approved mutation in this editor flows through the
 * `CommandBus` (`agent.applyCommand` / `applyCommands`), which owns
 * the canonical undo/redo stack. Installing PM's history plugin
 * alongside the bus produced two parallel stacks: the toolbar's
 * `Undo` button toggled the bus, while `Cmd-Z` toggled PM, and the
 * snapshot-driven re-projection silently leaked into PM's history
 * as a giant `replaceWith` step that PM would later try to invert
 * — desyncing the bus from the doc.
 *
 * We bind `Mod-Z` / `Mod-Y` / `Mod-Shift-Z` directly to the bus so
 * there is exactly one history per document. See
 * `spec/shared/agent-api.md` for the invariant.
 */

export interface MountOptions {
  agent: DocxAgent;
  agentId?: string;
  source?: "human" | "agent" | "system";
  onUnsupported?: (events: UnsupportedTx[]) => void;
  onError?: (err: unknown) => void;
  /**
   * Optional ProseMirror plugins to install alongside the built-in
   * history / keymap stack. P3.3 / W11 uses this to inject the
   * page-decorations plugin from the React layer (where chunkIntoPages
   * has access to the live `DocxAgent`). The plugins MUST be read-only
   * — any plugin that dispatches transactions must route them through
   * the bus to keep the round-trip contract.
   */
  extraPlugins?: ReadonlyArray<Plugin>;
  /**
   * Editor interaction mode:
   *   - `"edit"` (default) — direct editing; PM steps translate
   *     to plain `insert-text` / `delete-range` commands.
   *   - `"suggest"` — typing and deletions are routed through
   *     `insert-text-tracked` / `delete-range-tracked`, producing
   *     `<w:ins>` / `<w:del>` revision wrappers (the
   *     "Suggesting" / "Track Changes" surface).
   *   - `"view"` — read-only; PM rejects all user input.
   */
  editMode?: "edit" | "suggest" | "view";
  /**
   * Author attribution for Suggesting mode. Required when
   * `editMode === "suggest"`; ignored otherwise.
   */
  trackedAuthor?: string;
}

export interface MountResult {
  view: EditorView;
  destroy: () => void;
  /**
   * Update the live edit mode without rebuilding the editor. The
   * mount creates this once and the React layer flips it whenever
   * the user picks a different mode in the toolbar. Returning a
   * mutator (rather than re-mounting) keeps the cursor, scroll
   * position, and selection intact across mode changes.
   */
  setEditMode: (mode: "edit" | "suggest" | "view", author?: string) => void;
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
      keymap(busUndoRedoKeymap(agent)),
      keymap(baseKeymap),
      ...(opts.extraPlugins ?? []),
    ],
  });

  // Counts mutations that the funnel has dispatched into the bus and is
  // waiting to "echo" through the subscribe callback. We use a counter
  // (not a boolean) because handlers run synchronously inside
  // `applyCommands`, so subscribe fires once per command.
  let pendingFunnelCount = 0;
  let isProjecting = false;
  // When the funnel routes input through the tracked-changes commands
  // (`<w:ins>`/`<w:del>` wrappers) we deliberately skip PM's optimistic
  // apply and let the snapshot-driven re-projection paint the
  // `revision_mark` spans. This holds the desired post-input selection
  // (computed from the pre-input doc plus the change delta) so the
  // re-projection can clamp to the correct cursor instead of falling
  // back to the stale optimistic position.
  let pendingTrackedSelection: number | null = null;

  // Live-mutable mode + author state. Captured by closure so
  // `dispatchTransaction` always sees the latest values without
  // re-mounting the view (which would clobber selection / scroll).
  let editMode: "edit" | "suggest" | "view" = opts.editMode ?? "edit";
  let trackedAuthor: string = opts.trackedAuthor ?? "";

  const view = new EditorView(target, {
    state,
    editable: () => editMode !== "view",
    dispatchTransaction(tx: Transaction) {
      // Projections initiated by the bus go through unmodified.
      if (tx.getMeta("from-bus") === true || isProjecting) {
        view.updateState(view.state.apply(tx));
        return;
      }

      const before = view.state;

      if (tx.steps.length === 0) {
        // Selection-only / metadata transactions: apply locally and
        // return. There's nothing to mirror to the bus.
        view.updateState(before.apply(tx));
        return;
      }

      const result = transactionToCommands(tx, before, {
        ...(opts.agentId !== undefined ? { agentId: opts.agentId } : {}),
        source: opts.source ?? "human",
        mode: editMode === "view" ? "edit" : editMode,
        author: trackedAuthor,
      });
      if (result.unsupported.length > 0) opts.onUnsupported?.(result.unsupported);

      // Drift guard: if the funnel can't translate ANY of the
      // transaction's steps into commands, refuse to apply the
      // transaction locally. The previous behaviour was to
      // optimistically `view.updateState(before.apply(tx))` and
      // then fall through with an empty `result.commands`, which
      // mutated PM's doc without a corresponding bus mutation.
      // The doc and the snapshot would silently diverge — undo
      // couldn't reach the change, and the next from-bus
      // projection would erase whatever the user had just
      // produced. Rejecting up-front keeps the invariant
      // "PM's doc == bus snapshot projection" intact, and
      // `onUnsupported` already fired above so the host UI can
      // toast (e.g. "this list reflow isn't supported yet").
      if (result.commands.length === 0) {
        return;
      }

      const isTracked = result.commands.some(
        (c) => c.type === "docx:insert-text-tracked" || c.type === "docx:delete-range-tracked"
      );

      if (isTracked) {
        // Tracked-changes path: PM's optimistic apply would land the
        // user's text as PLAIN inline content (no `revision_mark`).
        // We deliberately skip the optimistic apply and instead let
        // the bus subscribe callback re-project the snapshot — that
        // pulls the `<w:ins>`/`<w:del>` wrappers (rendered as
        // `pm-revision-{ins,del}` spans) into the DOM so the
        // underline + margin balloons surface immediately.
        //
        // We DO need to compute where the cursor should land after
        // re-projection. PM's selection in `before.apply(tx)` already
        // reflects the post-input cursor in a hypothetical "plain
        // text" world; that maps 1:1 onto the snapshot doc because
        // `revision_mark` is a Mark (not a Node) and adds no
        // structural offset. We stash that target so the subscribe
        // re-projection clamps to it instead of the stale `before`
        // selection.
        const optimisticState = before.apply(tx);
        pendingTrackedSelection = optimisticState.selection.from;
        // Don't call view.updateState — we want PM's DOM to wait for
        // the snapshot-driven re-projection so the marks land in one
        // step. Otherwise the user briefly sees plain text, then the
        // re-projection swaps it for the marked variant.
        void agent.applyCommands(result.commands).catch((err) => {
          pendingTrackedSelection = null;
          opts.onError?.(err);
        });
        return;
      }

      // Edit mode (the historical fast path): apply optimistically
      // and mark the resulting bus mutations as "self-originated"
      // so the subscribe callback does not echo them back as a
      // re-projection (which would clobber the selection we just
      // preserved).
      view.updateState(before.apply(tx));
      pendingFunnelCount += result.commands.length;
      void agent.applyCommands(result.commands).catch((err) => {
        // Make sure we don't leak the suppression count if the
        // bus rejected our commands.
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
    // the host UI, or — most commonly — a tracked-changes mutation that
    // dispatchTransaction routed through the bus without an optimistic
    // PM apply). Re-project the snapshot and map the selection through
    // the structural change so the cursor doesn't jump to position 0.
    //
    // For the tracked-changes path, dispatchTransaction stashed the
    // intended post-input cursor in `pendingTrackedSelection`; we
    // consume that here so the cursor lands AFTER the inserted text
    // (or at the start of a tracked deletion) instead of at the
    // pre-input position.
    isProjecting = true;
    try {
      const trackedTarget = pendingTrackedSelection;
      pendingTrackedSelection = null;
      const desiredAnchor = trackedTarget ?? view.state.selection.from;
      const desiredHead = trackedTarget ?? view.state.selection.to;
      const oldDocSize = view.state.doc.content.size;
      const pm = docToPM(agent.getSnapshot());
      const tr = view.state.tr.replaceWith(0, oldDocSize, pm.content);
      const docSize = tr.doc.content.size;
      const anchor = clampSelection(tr.doc, Math.min(desiredAnchor, docSize));
      const head = clampSelection(tr.doc, Math.min(desiredHead, docSize));
      try {
        tr.setSelection(TextSelection.create(tr.doc, anchor, head));
      } catch {
        /* selection couldn't be set; PM will fall back to default */
      }
      tr.setMeta("from-bus", true);
      // Belt-and-suspenders against any future plugin (collab,
      // list keymap, alternate history, etc.) that might inspect
      // `tx.docChanged` and decide to record this as an
      // undoable user edit. The bus already owns the canonical
      // history; a from-bus projection IS that canonical state
      // being painted, not a new edit. Marking it as
      // `addToHistory: false` keeps the contract — "PM is a view,
      // the bus is the model" — even if someone re-installs a
      // history plugin upstream.
      tr.setMeta("addToHistory", false);
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
    setEditMode(nextMode, nextAuthor) {
      editMode = nextMode;
      if (typeof nextAuthor === "string") trackedAuthor = nextAuthor;
      // Trigger PM to recompute `editable()` (which it caches per
      // dispatch). Pushing an empty meta-only transaction is the
      // canonical "force a re-render without changing the doc" hack.
      const tr = view.state.tr.setMeta("from-bus", true);
      tr.setMeta("addToHistory", false);
      view.dispatch(tr);
    },
  };
}

/**
 * Build a `prosemirror-keymap` map that routes Mod-Z / Mod-Y /
 * Mod-Shift-Z directly to the bus. Returning `true` from each
 * binding consumes the event so the browser doesn't ALSO try to
 * undo (which on a contenteditable surface produces a phantom
 * "execCommand" undo that bypasses both PM and the bus and was a
 * known source of phantom edits before the bus existed).
 *
 * The bindings are written so that pressing the chord with nothing
 * to undo / redo still consumes the event — same as the inline
 * handlers in the XLSX and PPTX editors. That's deliberate: a
 * disabled toolbar button is a better signal to the user than
 * silently letting the browser pop up its own undo on the live
 * document surface.
 */
function busUndoRedoKeymap(agent: DocxAgent): Record<
  string,
  (state: EditorState, dispatch?: (tr: Transaction) => void, view?: EditorView) => boolean
> {
  const undo = (): boolean => {
    if (agent.canUndo()) agent.undo();
    return true;
  };
  const redo = (): boolean => {
    if (agent.canRedo()) agent.redo();
    return true;
  };
  return {
    "Mod-z": undo,
    "Mod-y": redo,
    "Mod-Shift-z": redo,
  };
}

function clampSelection(doc: import("prosemirror-model").Node, pos: number): number {
  const max = doc.content.size;
  const p = Math.max(0, Math.min(max, pos));
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
