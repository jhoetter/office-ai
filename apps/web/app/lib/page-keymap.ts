import { Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import type { DocxAgent } from "@officeai/docx";
import { currentParagraphId, pmSelectionToRange } from "./format-helpers";
import { movePageRelative } from "./page-decorations";

/**
 * P3.5 — page-aware keymap plugin.
 *
 * Binds:
 *  - `Mod-Enter` (Ctrl+Enter on Linux/Win, Cmd+Enter on macOS) →
 *    dispatch `docx:insert-page-break` against the paragraph + offset
 *    that the caret is in. Mirrors Word's "Insert → Page Break"
 *    shortcut.
 *  - `PageDown` → advance the caret by one page chunk (W21).
 *  - `PageUp` → retreat the caret by one page chunk (W21).
 *
 * All commands fall through to PM's default behavior when the
 * targeted action is not applicable (no paragraph at the caret,
 * caret already on the first/last page, etc.) so users never lose
 * a keystroke.
 *
 * Implemented via `props.handleKeyDown` rather than `prosemirror-keymap`
 * so the editor host (apps/web) doesn't need to add a direct dep on the
 * keymap package; the docx renderer already pulls it in for its own
 * default keymap.
 */
export function pageKeymapPlugin(agent: DocxAgent): Plugin {
  return new Plugin({
    props: {
      handleKeyDown(view: EditorView, event: KeyboardEvent): boolean {
        const isMod = event.metaKey || event.ctrlKey;
        if (event.key === "Enter" && isMod && !event.shiftKey && !event.altKey) {
          return handleInsertPageBreak(view, agent);
        }
        if (event.key === "PageDown") {
          return movePageRelative(view, 1);
        }
        if (event.key === "PageUp") {
          return movePageRelative(view, -1);
        }
        return false;
      },
    },
  });
}

function handleInsertPageBreak(view: EditorView, agent: DocxAgent): boolean {
  const paragraphId = currentParagraphId(view.state);
  if (!paragraphId) return false;
  const range = pmSelectionToRange(view.state);
  const offset = range.start.offset;
  void agent.applyCommand({
    type: "docx:insert-page-break",
    payload: { paragraphId, offset },
    source: "human",
  });
  return true;
}
