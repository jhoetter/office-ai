import { Plugin, PluginKey, type EditorState, type Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { Decoration, DecorationSet } from "prosemirror-view";

/**
 * "Show formatting marks" plugin (Word's pilcrow ¶ toggle).
 *
 * When enabled, overlays:
 *   - `¶` widget at the end of every paragraph,
 *   - `·` (middle-dot) glyphs over every space character,
 *   - `→` widget for `tab` nodes,
 *   - `↵` widget for `hard_break` nodes.
 *
 * Page-break dividers are already drawn by `pageDecorationsPlugin`,
 * so this plugin deliberately leaves them alone.
 *
 * Toggling dispatches a meta-only transaction tagged with
 * {@link FORMATTING_MARKS_META}; no doc mutation occurs, so the bus
 * is never touched and the round-trip contract is preserved.
 *
 * Persistence: the toggle lives only in plugin state for the current
 * mount. A future iteration could mirror it into a user preference
 * (localStorage) — out of scope for now.
 */
export const FORMATTING_MARKS_KEY = new PluginKey<FormattingMarksState>("formattingMarks");
export const FORMATTING_MARKS_META = "formatting-marks-toggle";

export interface FormattingMarksState {
  readonly enabled: boolean;
  readonly decorations: DecorationSet;
}

export function formattingMarksPlugin(): Plugin<FormattingMarksState> {
  return new Plugin<FormattingMarksState>({
    key: FORMATTING_MARKS_KEY,
    state: {
      init(): FormattingMarksState {
        return { enabled: false, decorations: DecorationSet.empty };
      },
      apply(tr: Transaction, prev: FormattingMarksState, _oldState, newState): FormattingMarksState {
        const meta = tr.getMeta(FORMATTING_MARKS_META);
        let enabled = prev.enabled;
        if (meta && typeof meta === "object" && "enabled" in meta) {
          enabled = Boolean((meta as { enabled: unknown }).enabled);
        }
        if (!enabled) {
          if (prev.decorations === DecorationSet.empty && enabled === prev.enabled) return prev;
          return { enabled: false, decorations: DecorationSet.empty };
        }
        // Recompute decorations whenever the doc structure changes
        // OR when the toggle just flipped on. Cheap walk: O(n) over
        // the doc, no per-keystroke cost when disabled.
        if (!tr.docChanged && enabled === prev.enabled) return prev;
        return { enabled: true, decorations: buildDecorations(newState) };
      },
    },
    props: {
      decorations(state) {
        const s = FORMATTING_MARKS_KEY.getState(state);
        return s ? s.decorations : null;
      },
    },
  });
}

function buildDecorations(state: EditorState): DecorationSet {
  const decos: Decoration[] = [];
  state.doc.descendants((node, pos) => {
    if (node.type.name === "paragraph") {
      // Pilcrow at the end of the paragraph (just before the closing
      // node boundary). `nodeSize - 1` lands on the position right
      // after the last inline child, which is where Word draws ¶.
      const end = pos + node.nodeSize - 1;
      decos.push(
        Decoration.widget(end, () => makePilcrow(), { side: 1, key: `pilcrow-${pos}` })
      );
      return true;
    }
    if (node.type.name === "tab") {
      decos.push(
        Decoration.widget(pos, () => makeTabArrow(), { side: -1, key: `tab-${pos}` })
      );
      return false;
    }
    if (node.type.name === "hard_break") {
      decos.push(
        Decoration.widget(pos, () => makeBreakArrow(), { side: -1, key: `br-${pos}` })
      );
      return false;
    }
    if (node.isText) {
      const text = node.text ?? "";
      if (!text.includes(" ")) return false;
      // One inline decoration per contiguous run of spaces; keeps
      // the DecorationSet small even for paragraphs with lots of
      // spacing.
      let i = 0;
      while (i < text.length) {
        if (text[i] === " ") {
          let j = i + 1;
          while (j < text.length && text[j] === " ") j++;
          decos.push(
            Decoration.inline(pos + i, pos + j, { class: "fmt-mark-space" })
          );
          i = j;
        } else {
          i++;
        }
      }
      return false;
    }
    return true;
  });
  return DecorationSet.create(state.doc, decos);
}

function makePilcrow(): HTMLElement {
  const el = document.createElement("span");
  el.className = "fmt-mark-pilcrow";
  el.setAttribute("aria-hidden", "true");
  el.textContent = "\u00B6";
  return el;
}

function makeTabArrow(): HTMLElement {
  const el = document.createElement("span");
  el.className = "fmt-mark-tab";
  el.setAttribute("aria-hidden", "true");
  el.textContent = "\u2192";
  return el;
}

function makeBreakArrow(): HTMLElement {
  const el = document.createElement("span");
  el.className = "fmt-mark-break";
  el.setAttribute("aria-hidden", "true");
  el.textContent = "\u21B5";
  return el;
}

/** Read the current on/off state from any EditorState. */
export function isFormattingMarksOn(state: EditorState): boolean {
  return FORMATTING_MARKS_KEY.getState(state)?.enabled ?? false;
}

/** Imperatively toggle the plugin. Dispatches a meta-only tx. */
export function toggleFormattingMarks(view: EditorView): boolean {
  const next = !isFormattingMarksOn(view.state);
  view.dispatch(view.state.tr.setMeta(FORMATTING_MARKS_META, { enabled: next }));
  return next;
}

/** Imperatively set the on/off state (idempotent). */
export function setFormattingMarks(view: EditorView, enabled: boolean): void {
  if (isFormattingMarksOn(view.state) === enabled) return;
  view.dispatch(view.state.tr.setMeta(FORMATTING_MARKS_META, { enabled }));
}
