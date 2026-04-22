import { Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import type { DocxAgent, DocxSnapshot } from "@officeai/docx";
import {
  activeMarks,
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
 *   Paragraph alignment (Google-Docs parity — `Mod+Shift+L/E/R/J`
 *   instead of Word's `Mod+L/E/R/J` so the bare modifier chords stay
 *   with the browser: `Cmd+R` reloads the tab, `Cmd+L` focuses the
 *   address bar, etc. Word's bare-modifier alignment was a frequent
 *   "the page won't refresh!" complaint):
 *     Mod-Shift-L / Mod-Shift-E / Mod-Shift-R / Mod-Shift-J →
 *       left / center / right / justify
 *
 *   Line spacing:
 *     Mod-1 / Mod-5 / Mod-2  → single / 1.5 / double
 *
 *   Indent:
 *     Mod-M                  → +360 twips (¼")
 *     Mod-Shift-M            → −360 twips
 *
 *   Lists (Google-Docs parity — bullet on Shift+8, numbered on
 *   Shift+7, matching the digit's typewriter glyph):
 *     Mod-Shift-8            → toggle bullet
 *     Mod-Shift-7            → toggle numbered
 *
 *   View:
 *     Mod-Shift-P            → toggle "Show formatting marks" (¶)
 *                              (moved off Mod-Shift-8 so bullet list
 *                              can take the Google-Docs slot)
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
  // Tab / Shift+Tab inside list items demote/promote the bullet level
  // (Word parity). We handle these BEFORE the Mod gate because Tab
  // doesn't carry a modifier; PM falls back to its default Tab handler
  // (insert literal tab) if the caret isn't in a list paragraph.
  if (event.key === "Tab" && !event.metaKey && !event.ctrlKey && !event.altKey) {
    const adjusted = adjustListLevel(view, agent, event.shiftKey ? -1 : +1);
    if (adjusted) return true;
  }

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
  // Google-Docs chord set: `Mod+Shift+L/E/R/J`. We deliberately do
  // NOT bind the bare-modifier variants Word uses, so that on every
  // major browser:
  //   - `Cmd/Ctrl+R` keeps reloading the page (the most-missed one),
  //   - `Cmd/Ctrl+L` keeps focusing the address bar,
  //   - `Cmd/Ctrl+E` keeps doing the browser's own thing,
  //   - `Cmd/Ctrl+J` keeps opening the downloads pane.
  // Resolved via `event.code` (`KeyL/E/R/J`) so the lookup stays
  // layout-independent — `event.key` is the Shift-mapped character
  // on most keyboards (e.g. `L` instead of `l`) and depending on
  // the layout we'd otherwise need both branches.
  if (shift && !alt) {
    if (code === "KeyL") {
      return setAlignment(view, agent, "left");
    }
    if (code === "KeyE") {
      return setAlignment(view, agent, "center");
    }
    if (code === "KeyR") {
      return setAlignment(view, agent, "right");
    }
    if (code === "KeyJ") {
      return setAlignment(view, agent, "justify");
    }
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
  // Google-Docs convention: bullet = Shift+8 (the * glyph), numbered
  // = Shift+7 (digit). Bullet list moved off `Mod+Shift+L` so the
  // alignment block above can take that slot.
  if (shift && !alt) {
    if (code === "Digit8") {
      return toggleList(view, agent, "bullet");
    }
    if (code === "Digit7") {
      return toggleList(view, agent, "ordered");
    }
  }

  // View: pilcrow / "show formatting marks" -----------------------
  // Was on Mod+Shift+8 (Word's documented chord) but that slot now
  // belongs to bullet list (Google-Docs parity). Re-homed onto
  // Mod+Shift+P, which is unbound by every other surface in the
  // app. The toolbar pilcrow button is unchanged so users who
  // never learn the new chord still have a one-click affordance.
  if (shift && !alt && code === "KeyP") {
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

/**
 * Map our `format-range` payload mark names onto the PM mark
 * type names used by `activeMarks`. The schema spells
 * strikethrough as `strikethrough`, but the OOXML payload field is
 * `strike`; the rest line up directly.
 */
function pmMarkName(mark: ToggleableMark): string {
  return mark === "strike" ? "strikethrough" : mark;
}

function toggleMark(view: EditorView, agent: DocxAgent, mark: ToggleableMark): boolean {
  if (view.state.selection.empty) return false;
  const range = pmSelectionToRange(view.state);
  // Use PM's mark set for the selection (Word-style "any text in the
  // selection has the mark" → toggle off; otherwise → toggle on).
  // Probing the snapshot by `(paragraph, offset)` would mis-read the
  // boundary case where the selection start sits exactly between an
  // unmarked and a marked run — `activeMarks` only inspects text
  // nodes that fall *within* the selection, so it doesn't suffer
  // from that off-by-one.
  const marks = activeMarks(view.state);
  const currentlyOn = marks.has(pmMarkName(mark));
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

function findParagraphById(
  snap: DocxSnapshot,
  paragraphId: string
): { properties: { numbering?: { numId: number; ilvl: number } } } | null {
  for (const block of snap.root.body) {
    if (block.kind === "paragraph" && block.id === paragraphId) return block;
  }
  return null;
}

/**
 * Word-parity Tab / Shift+Tab demote/promote inside lists.
 *
 * Returns `true` and dispatches `docx:set-paragraph-list` (preserving
 * the existing numId, only bumping ilvl) when:
 *   - the caret sits in a list paragraph
 *   - the resulting level stays in `[0, 8]` (Word's hard cap on
 *     `<w:ilvl>` per OOXML's `numFmt` definitions)
 *
 * Returns `false` (and does NOT dispatch) when:
 *   - the caret isn't in a list paragraph (PM's default Tab inserts
 *     a literal tab character — we leave that contract intact),
 *   - the level is already at the relevant boundary (so a stranded
 *     Shift+Tab on a top-level bullet doesn't silently no-op the
 *     Tab key and trap the user with no escape).
 *
 * The `direction` argument is `+1` for demote (Tab) and `-1` for
 * promote (Shift+Tab), matching the visual mental model of "Tab
 * pushes deeper".
 */
function adjustListLevel(view: EditorView, agent: DocxAgent, direction: 1 | -1): boolean {
  const paragraphId = currentParagraphId(view.state);
  if (!paragraphId) return false;
  const snap = agent.getSnapshot();
  const para = findParagraphById(snap, paragraphId);
  const numbering = para?.properties.numbering;
  if (!numbering) return false;
  const nextIlvl = numbering.ilvl + direction;
  // Word caps `<w:ilvl>` at 8 (levels 0..8 — the visible nine
  // indents in the bullets ribbon dropdown). Below 0 there's no
  // such thing as a "negative" indent.
  if (nextIlvl < 0 || nextIlvl > 8) return false;
  void agent.applyCommand({
    type: "docx:set-paragraph-list",
    payload: { paragraphId, numId: numbering.numId, ilvl: nextIlvl },
    source: "human",
  });
  return true;
}
