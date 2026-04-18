/**
 * Per-product keyboard-shortcut catalog. Source-of-truth for the
 * {@link KeyboardShortcutsDialog} so the help panel always reflects
 * what the editor actually does.
 *
 * Cross-references:
 *   - DOCX: `apps/web/app/lib/word-shortcuts-keymap.ts` (inline marks,
 *     paragraph + list shortcuts), `apps/web/app/lib/page-keymap.ts`
 *     (Mod+Enter page break), `packages/docx/src/renderer/mount.ts`
 *     (PM history covers `Mod+Z`/`Mod+Y`/`Mod+Shift+Z`).
 *   - XLSX: `apps/web/app/xlsx-editor/XlsxEditor.tsx` `onSurfaceKeyDown`.
 *   - PPTX: `apps/web/app/lib/pptx-shortcuts.ts`.
 *
 * Anything tagged `planned` renders greyed-out with a "soon" pill so
 * the dialog stays honest about what is and isn't wired.
 */
export type ShortcutKey =
  | "Mod"
  | "Shift"
  | "Alt"
  | "Enter"
  | "Tab"
  | "Esc"
  | "Backspace"
  | "Delete"
  | "Space"
  | "ArrowUp"
  | "ArrowDown"
  | "ArrowLeft"
  | "ArrowRight"
  | "Home"
  | "End"
  | "PageUp"
  | "PageDown"
  | "F2"
  | "F4"
  | (string & { __plain?: never });

export type ShortcutProduct = "docx" | "xlsx" | "pptx";

export interface ShortcutEntry {
  /** Tokens rendered as individual `<kbd>` chips. */
  readonly keys: ReadonlyArray<ShortcutKey>;
  readonly label: string;
  readonly category: string;
  readonly status?: "implemented" | "planned";
}

export interface ShortcutCatalog {
  readonly product: ShortcutProduct;
  readonly title: string;
  readonly entries: ReadonlyArray<ShortcutEntry>;
}

// ─── DOCX ────────────────────────────────────────────────────────────
// Reflects `wordShortcutsKeymapPlugin` + `pageKeymapPlugin` + PM
// `baseKeymap`. Keep entries in this order; the dialog groups by
// `category` and preserves first-seen order within a group.
export const DOCX_SHORTCUTS: ReadonlyArray<ShortcutEntry> = [
  // Inline marks
  { keys: ["Mod", "B"], label: "Bold", category: "Inline marks" },
  { keys: ["Mod", "I"], label: "Italic", category: "Inline marks" },
  { keys: ["Mod", "U"], label: "Underline", category: "Inline marks" },
  { keys: ["Mod", "Shift", "X"], label: "Strikethrough", category: "Inline marks" },

  // Paragraph alignment
  { keys: ["Mod", "L"], label: "Align left", category: "Paragraph" },
  { keys: ["Mod", "E"], label: "Align center", category: "Paragraph" },
  { keys: ["Mod", "R"], label: "Align right", category: "Paragraph" },
  { keys: ["Mod", "J"], label: "Justify", category: "Paragraph" },

  // Line spacing
  { keys: ["Mod", "1"], label: "Single line spacing", category: "Paragraph" },
  { keys: ["Mod", "5"], label: "1.5 line spacing", category: "Paragraph" },
  { keys: ["Mod", "2"], label: "Double line spacing", category: "Paragraph" },

  // Indent
  { keys: ["Mod", "M"], label: "Increase indent", category: "Paragraph" },
  { keys: ["Mod", "Shift", "M"], label: "Decrease indent", category: "Paragraph" },

  // Lists
  { keys: ["Mod", "Shift", "L"], label: "Toggle bullet list", category: "Lists" },
  { keys: ["Mod", "Shift", "7"], label: "Toggle numbered list", category: "Lists" },

  // View
  { keys: ["Mod", "Shift", "8"], label: "Show formatting marks (¶)", category: "View" },

  // Document structure
  { keys: ["Mod", "Enter"], label: "Insert page break", category: "Document" },
  { keys: ["Mod", "Shift", "Enter"], label: "Insert section break (next page)", category: "Document" },

  // Collaboration
  { keys: ["Mod", "K"], label: "Insert hyperlink", category: "Collaboration" },
  { keys: ["Mod", "Alt", "M"], label: "Add comment", category: "Collaboration" },

  // History
  { keys: ["Mod", "Z"], label: "Undo", category: "History" },
  { keys: ["Mod", "Shift", "Z"], label: "Redo", category: "History" },

  // Help
  { keys: ["Mod", "/"], label: "Show keyboard shortcuts", category: "Help" },
];

// ─── XLSX ────────────────────────────────────────────────────────────
// Reflects `XlsxEditor.tsx`'s `onSurfaceKeyDown`. New B/I/U + numfmt
// shortcuts are added in the same change as this catalog; treat them
// as implemented once `xlsx-shortcuts` ships in the same commit.
export const XLSX_SHORTCUTS: ReadonlyArray<ShortcutEntry> = [
  // Navigation
  { keys: ["ArrowUp"], label: "Move selection up", category: "Navigation" },
  { keys: ["ArrowDown"], label: "Move selection down", category: "Navigation" },
  { keys: ["ArrowLeft"], label: "Move selection left", category: "Navigation" },
  { keys: ["ArrowRight"], label: "Move selection right", category: "Navigation" },
  { keys: ["Mod", "ArrowUp"], label: "Jump to data edge (up)", category: "Navigation" },
  { keys: ["Mod", "ArrowDown"], label: "Jump to data edge (down)", category: "Navigation" },
  { keys: ["Mod", "ArrowLeft"], label: "Jump to data edge (left)", category: "Navigation" },
  { keys: ["Mod", "ArrowRight"], label: "Jump to data edge (right)", category: "Navigation" },
  { keys: ["Shift", "ArrowUp"], label: "Extend selection up", category: "Navigation" },
  { keys: ["Shift", "ArrowDown"], label: "Extend selection down", category: "Navigation" },
  { keys: ["Shift", "ArrowLeft"], label: "Extend selection left", category: "Navigation" },
  { keys: ["Shift", "ArrowRight"], label: "Extend selection right", category: "Navigation" },
  { keys: ["Home"], label: "Jump to start of row", category: "Navigation" },
  { keys: ["Mod", "Home"], label: "Jump to A1", category: "Navigation" },
  { keys: ["Mod", "End"], label: "Jump to last used cell", category: "Navigation" },
  { keys: ["Tab"], label: "Move selection right (commit)", category: "Navigation" },
  { keys: ["Enter"], label: "Move selection down (commit)", category: "Navigation" },

  // Editing
  { keys: ["F2"], label: "Edit active cell in formula bar", category: "Editing" },
  { keys: ["Esc"], label: "Cancel selection / dismiss clipboard", category: "Editing" },
  { keys: ["Backspace"], label: "Clear selection", category: "Editing" },
  { keys: ["Delete"], label: "Clear selection / delete row or column", category: "Editing" },

  // Inline marks
  { keys: ["Mod", "B"], label: "Bold", category: "Inline marks" },
  { keys: ["Mod", "I"], label: "Italic", category: "Inline marks" },
  { keys: ["Mod", "U"], label: "Underline", category: "Inline marks" },

  // Number formats
  { keys: ["Mod", "Shift", "1"], label: "Number format (#,##0.00)", category: "Number format" },
  { keys: ["Mod", "Shift", "4"], label: "Currency format ($#,##0.00)", category: "Number format" },
  { keys: ["Mod", "Shift", "5"], label: "Percent format (0%)", category: "Number format" },
  { keys: ["F4"], label: "Repeat last format", category: "Number format", status: "planned" },

  // History
  { keys: ["Mod", "Z"], label: "Undo", category: "History" },
  { keys: ["Mod", "Shift", "Z"], label: "Redo", category: "History" },
  { keys: ["Mod", "Y"], label: "Redo (alt)", category: "History" },

  // Help
  { keys: ["Mod", "/"], label: "Show keyboard shortcuts", category: "Help" },
];

// ─── PPTX ────────────────────────────────────────────────────────────
// Reflects `usePptxShortcuts`. `Mod+D` (duplicate shape) has no
// matching command yet, so it stays `planned`.
export const PPTX_SHORTCUTS: ReadonlyArray<ShortcutEntry> = [
  // Selection
  { keys: ["Backspace"], label: "Delete selected shape(s)", category: "Selection" },
  { keys: ["Delete"], label: "Delete selected shape(s)", category: "Selection" },

  // Move
  { keys: ["ArrowUp"], label: "Nudge shape up (1px)", category: "Move" },
  { keys: ["ArrowDown"], label: "Nudge shape down (1px)", category: "Move" },
  { keys: ["ArrowLeft"], label: "Nudge shape left (1px)", category: "Move" },
  { keys: ["ArrowRight"], label: "Nudge shape right (1px)", category: "Move" },
  { keys: ["Shift", "ArrowUp"], label: "Move shape up (10px)", category: "Move" },
  { keys: ["Shift", "ArrowDown"], label: "Move shape down (10px)", category: "Move" },
  { keys: ["Shift", "ArrowLeft"], label: "Move shape left (10px)", category: "Move" },
  { keys: ["Shift", "ArrowRight"], label: "Move shape right (10px)", category: "Move" },

  // Inline marks (selected text frame)
  { keys: ["Mod", "B"], label: "Bold", category: "Inline marks" },
  { keys: ["Mod", "I"], label: "Italic", category: "Inline marks" },
  { keys: ["Mod", "U"], label: "Underline", category: "Inline marks" },

  // Slide
  { keys: ["Mod", "M"], label: "Add new slide", category: "Slide" },
  { keys: ["Mod", "Shift", "D"], label: "Duplicate current slide", category: "Slide" },
  { keys: ["PageUp"], label: "Previous slide", category: "Slide" },
  { keys: ["PageDown"], label: "Next slide", category: "Slide" },
  { keys: ["Mod", "D"], label: "Duplicate selected shape", category: "Selection", status: "planned" },

  // History
  { keys: ["Mod", "Z"], label: "Undo", category: "History" },
  { keys: ["Mod", "Shift", "Z"], label: "Redo", category: "History" },

  // Help
  { keys: ["Mod", "/"], label: "Show keyboard shortcuts", category: "Help" },
];

export const SHORTCUT_CATALOGS: Readonly<Record<ShortcutProduct, ShortcutCatalog>> = {
  docx: {
    product: "docx",
    title: "DOCX editor",
    entries: DOCX_SHORTCUTS,
  },
  xlsx: {
    product: "xlsx",
    title: "XLSX editor",
    entries: XLSX_SHORTCUTS,
  },
  pptx: {
    product: "pptx",
    title: "PPTX editor",
    entries: PPTX_SHORTCUTS,
  },
};
