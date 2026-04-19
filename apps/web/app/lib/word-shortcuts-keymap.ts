import { Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import type { DocxAgent, DocxSnapshot } from "@officeai/docx";
import {
  currentParagraphId,
  currentParagraphIndex,
  discoverNumId,
  pmSelectionToRange,
} from "./format-helpers";

/**
 * Word-parity keyboard shortcut plugin.
 *
 * Binds the "muscle-memory" subset of Microsoft Word shortcuts that
 * map cleanly onto a registered command in `@officeai/docx`. Pure
 * keyboard sugar — every action ultimately dispatches via
 * `agent.applyCommand`, so the bus stays the source of truth and
 * tracked changes / undo / subscribe semantics are unchanged.
 *
 * Bindings (Ctrl on Win/Linux, Cmd on macOS):
 *
 *   Inline marks:
 *     Mod-B                  → toggle bold
 *     Mod-I                  → toggle italic
 *     Mod-U                  → toggle underline
 *     Mod-Shift-X            → toggle strikethrough
 *
 *   Paragraph alignment:
 *     Mod-L / Mod-E / Mod-R / Mod-J → left / center / right / justify
 *
 *   Line spacing:
 *     Mod-1 / Mod-5 / Mod-2  → single / 1.5 / double
 *
 *   Indent:
 *     Mod-M                  → +360 twips (¼")
 *     Mod-Shift-M            → −360 twips
 *
 *   Lists:
 *     Mod-Shift-L            → toggle bullet
 *     Mod-Shift-7            → toggle numbered
 *
 *   View:
 *     Mod-Shift-8            → toggle "Show formatting marks" (¶)
 *
 *   Document structure:
 *     Mod-Enter              → page break  (handled by page-keymap.ts)
 *     Mod-Shift-Enter        → section break (next page)
 *
 *   Collaboration:
 *     Mod-K                  → insert hyperlink (host opens prompt)
 *     Mod-Alt-M              → add comment    (host opens composer)
 *
 * The host (`DocxEditor.tsx`) listens for two CustomEvents fired by
 * the plugin so the React layer owns the modal UI for hyperlink and
 * comment composition (the keymap can't open prompts without
 * leaking dependencies into the renderer):
 *
 *   - `docx:shortcut-insert-hyperlink`
 *   - `docx:shortcut-add-comment`
 *   - `docx:shortcut-toggle-formatting-marks`
 *
 * Every handler returns `false` from PM whenever the action can't
 * apply (no caret, empty selection where one is required, etc.) so
 * PM falls through to its default behaviour and no keystroke is
 * dropped on the floor.
 */
export const SHORTCUT_INSERT_HYPERLINK_EVENT = "docx:shortcut-insert-hyperlink";
export const SHORTCUT_ADD_COMMENT_EVENT = "docx:shortcut-add-comment";
export const SHORTCUT_TOGGLE_FORMATTING_MARKS_EVENT = "docx:shortcut-toggle-formatting-marks";

export interface InsertHyperlinkDetail {
  paragraphId: string;
  range: { start: number; end: number };
  selectionText: string;
}

export function wordShortcutsKeymapPlugin(agent: DocxAgent): Plugin {
  return new Plugin({
    props: {
      handleKeyDown(view: EditorView, event: KeyboardEvent): boolean {
        return dispatchShortcut(view, event, agent);
      },
    },
  });
}

/** Exposed for unit testing — pure function over (view, event, agent). */
export function dispatchShortcut(view: EditorView, event: KeyboardEvent, agent: DocxAgent): boolean {
  const isMod = event.metaKey || event.ctrlKey;
  if (!isMod) return false;

  const shift = event.shiftKey;
  const alt = event.altKey;
  // `event.code` is layout-independent (e.g. Digit7 always maps to
  // the physical "7" key, regardless of the Shift modifier producing
  // "&" on US layouts). `event.key` covers letters, where layout
  // already does the right thing.
  const key = event.key;
  const code = event.code;

  // Inline marks ------------------------------------------------------
  // PM's `baseKeymap` (registered in mountDocxEditor) does NOT include
  // Mod-B/Mod-I/Mod-U — only structural keys like Enter / Backspace /
  // Mod-Z. We bind the four core toggles ourselves so muscle memory
  // works out of the box.
  if ((key === "b" || key === "B") && !shift && !alt) {
    return toggleMark(view, agent, "bold");
  }
  if ((key === "i" || key === "I") && !shift && !alt) {
    return toggleMark(view, agent, "italic");
  }
  if ((key === "u" || key === "U") && !shift && !alt) {
    return toggleMark(view, agent, "underline");
  }
  if ((key === "x" || key === "X") && shift && !alt) {
    return toggleMark(view, agent, "strike");
  }

  // Alignment ---------------------------------------------------------
  if ((key === "l" || key === "L") && !shift && !alt) {
    return setAlignment(view, agent, "left");
  }
  if ((key === "e" || key === "E") && !shift && !alt) {
    return setAlignment(view, agent, "center");
  }
  if ((key === "r" || key === "R") && !shift && !alt) {
    return setAlignment(view, agent, "right");
  }
  if ((key === "j" || key === "J") && !shift && !alt) {
    return setAlignment(view, agent, "justify");
  }

  // Line spacing ------------------------------------------------------
  // Word's defaults: single = 240 twips, 1.5 = 360, double = 480.
  if (code === "Digit1" && !shift && !alt) {
    return setLineSpacing(view, agent, 240);
  }
  if (code === "Digit5" && !shift && !alt) {
    return setLineSpacing(view, agent, 360);
  }
  if (code === "Digit2" && !shift && !alt) {
    return setLineSpacing(view, agent, 480);
  }

  // Indent ------------------------------------------------------------
  if ((key === "m" || key === "M") && !alt) {
    return adjustIndent(view, agent, shift ? -360 : 360);
  }

  // Lists -------------------------------------------------------------
  if (shift && !alt) {
    if (code === "KeyL") {
      return toggleList(view, agent, "bullet");
    }
    if (code === "Digit7") {
      return toggleList(view, agent, "ordered");
    }
  }

  // View: pilcrow / "show formatting marks" -----------------------
  // Word's documented shortcut. Routed through a CustomEvent so the
  // React host can flip the plugin state and refresh its toolbar
  // pressed-state in one place.
  if (shift && !alt && code === "Digit8") {
    return requestToggleFormattingMarks(view);
  }

  // Document structure ----------------------------------------------
  if (event.key === "Enter" && shift && !alt) {
    return insertSectionBreak(view, agent);
  }

  // Collaboration ---------------------------------------------------
  if ((key === "k" || key === "K") && !shift && !alt) {
    return requestHyperlink(view);
  }
  if ((key === "m" || key === "M") && alt && !shift) {
    return requestComment(view);
  }

  return false;
}

type ToggleableMark = "bold" | "italic" | "underline" | "strike";

function toggleMark(view: EditorView, agent: DocxAgent, mark: ToggleableMark): boolean {
  if (view.state.selection.empty) return false;
  const range = pmSelectionToRange(view.state);
  // Best-effort "toggle": probe the run at the selection start; if the
  // mark is already on, turn it off, else on. We don't need a perfect
  // cross-run read here because the format-range handler is idempotent
  // and the toolbar's pressed-state derivation will refresh on
  // subscribe regardless.
  const snap = agent.getSnapshot();
  const currentlyOn = isMarkActiveAt(snap, range.start.paragraph, range.start.offset, mark);
  void agent.applyCommand({
    type: "docx:format-range",
    payload: {
      range,
      format: { [mark]: !currentlyOn } as {
        bold?: boolean;
        italic?: boolean;
        underline?: boolean;
        strike?: boolean;
      },
    },
    source: "human",
  });
  return true;
}

function setAlignment(
  view: EditorView,
  agent: DocxAgent,
  alignment: "left" | "center" | "right" | "justify"
): boolean {
  const paragraphId = currentParagraphId(view.state);
  if (!paragraphId) return false;
  void agent.applyCommand({
    type: "docx:set-paragraph-alignment",
    payload: { paragraphId, alignment },
    source: "human",
  });
  return true;
}

function setLineSpacing(view: EditorView, agent: DocxAgent, lineTwips: number): boolean {
  const paragraphId = currentParagraphId(view.state);
  if (!paragraphId) return false;
  void agent.applyCommand({
    type: "docx:set-paragraph-spacing",
    payload: { paragraphId, line: lineTwips, lineRule: "auto" },
    source: "human",
  });
  return true;
}

function adjustIndent(view: EditorView, agent: DocxAgent, deltaTwips: number): boolean {
  const paragraphId = currentParagraphId(view.state);
  if (!paragraphId) return false;
  void agent.applyCommand({
    type: "docx:set-paragraph-indent",
    payload: { paragraphId, deltaTwips },
    source: "human",
  });
  return true;
}

function toggleList(view: EditorView, agent: DocxAgent, kind: "bullet" | "ordered"): boolean {
  const paragraphId = currentParagraphId(view.state);
  if (!paragraphId) return false;
  const snap = agent.getSnapshot();
  const para = findParagraphById(snap, paragraphId);
  // If the caret paragraph is already a list item, strip the list
  // (toggle off). Otherwise, look up the abstract num for the
  // requested kind and apply.
  if (para?.properties.numbering) {
    void agent.applyCommand({
      type: "docx:remove-paragraph-list",
      payload: { paragraphId },
      source: "human",
    });
    return true;
  }
  const target = discoverNumId(snap, kind);
  if (!target) return false;
  void agent.applyCommand({
    type: "docx:set-paragraph-list",
    payload: { paragraphId, numId: target.numId, ilvl: target.ilvl },
    source: "human",
  });
  return true;
}

function insertSectionBreak(view: EditorView, agent: DocxAgent): boolean {
  const paragraphIndex = currentParagraphIndex(view.state);
  if (paragraphIndex < 0) return false;
  void agent.applyCommand({
    type: "docx:insert-section-break",
    payload: { paragraphIndex, type: "nextPage" },
    source: "human",
  });
  return true;
}

function requestHyperlink(view: EditorView): boolean {
  if (view.state.selection.empty) return false;
  const paragraphId = currentParagraphId(view.state);
  if (!paragraphId) return false;
  const range = pmSelectionToRange(view.state);
  // Only single-paragraph hyperlinks are supported by the model.
  if (range.start.paragraph !== range.end.paragraph) return false;
  const selectionText = view.state.doc.textBetween(
    view.state.selection.from,
    view.state.selection.to,
    " ",
    " "
  );
  const detail: InsertHyperlinkDetail = {
    paragraphId,
    range: { start: range.start.offset, end: range.end.offset },
    selectionText,
  };
  view.dom.dispatchEvent(new CustomEvent(SHORTCUT_INSERT_HYPERLINK_EVENT, { detail, bubbles: true }));
  return true;
}

function requestComment(view: EditorView): boolean {
  if (view.state.selection.empty) return false;
  view.dom.dispatchEvent(new CustomEvent(SHORTCUT_ADD_COMMENT_EVENT, { bubbles: true }));
  return true;
}

function requestToggleFormattingMarks(view: EditorView): boolean {
  view.dom.dispatchEvent(new CustomEvent(SHORTCUT_TOGGLE_FORMATTING_MARKS_EVENT, { bubbles: true }));
  return true;
}

/**
 * Probe whether the run that contains the given paragraph offset
 * already carries the requested mark. Conservative: when we can't
 * resolve the run cleanly (e.g. offset past the end), we report
 * "not active" so the toggle turns the mark on.
 */
function isMarkActiveAt(
  snap: DocxSnapshot,
  paragraphIndex: number,
  offset: number,
  mark: ToggleableMark
): boolean {
  const blocks = snap.root.body;
  let pIdx = -1;
  for (const b of blocks) {
    if (b.kind !== "paragraph") continue;
    pIdx++;
    if (pIdx !== paragraphIndex) continue;
    let cursor = 0;
    for (const child of b.children) {
      if (child.kind !== "run") continue;
      const text = runText(child);
      const next = cursor + text.length;
      if (offset >= cursor && offset <= next) {
        const rpr = child.properties ?? {};
        switch (mark) {
          case "bold":
            return Boolean(rpr.bold);
          case "italic":
            return Boolean(rpr.italic);
          case "underline":
            // `underline` is `boolean | string` in the model — both
            // truthy variants count as "on" for toggle purposes.
            return rpr.underline !== undefined && rpr.underline !== false;
          case "strike":
            return Boolean(rpr.strike);
        }
      }
      cursor = next;
    }
  }
  return false;
}

function runText(run: { children: ReadonlyArray<{ kind: string; text?: string }> }): string {
  let s = "";
  for (const c of run.children) {
    if (c.kind === "text" && typeof c.text === "string") s += c.text;
  }
  return s;
}

function findParagraphById(
  snap: DocxSnapshot,
  paragraphId: string
): { properties: { numbering?: { numId: number; ilvl: number } } } | null {
  for (const block of snap.root.body) {
    if (block.kind === "paragraph" && block.id === paragraphId) return block;
  }
  return null;
}
