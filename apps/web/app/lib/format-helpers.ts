import type { EditorState } from "prosemirror-state";
import type { DocxComment, DocxSnapshot, RevisionWrapper, InlineNode, Paragraph, Run } from "@officeai/docx";

/**
 * Pure helpers used by the editor toolbar / sidebars.
 *
 * These intentionally have no React or DOM dependencies so they can be
 * unit-tested in isolation. They translate between ProseMirror selection
 * positions, our DocxSnapshot model, and the small set of UI-state
 * derivations the toolbar needs (active style, active marks, etc.).
 */

export interface DocxRangePosition {
  paragraph: number;
  run: number;
  offset: number;
}

export interface DocxRange {
  start: DocxRangePosition;
  end: DocxRangePosition;
}

/**
 * Convert a ProseMirror absolute position into a paragraph index +
 * paragraph-wide offset. Run is fixed to 0 because format-range and
 * delete-range accept paragraph-wide offsets when run is undefined or 0
 * with our handler semantics (see packages/docx/src/commands/format-range.ts).
 */
export function pmPositionToDocx(state: EditorState, pos: number): DocxRangePosition {
  let paragraphIndex = -1;
  let result: DocxRangePosition | null = null;
  state.doc.descendants((node, nodePos) => {
    if (result) return false;
    if (node.type.name === "paragraph") {
      paragraphIndex++;
      if (pos >= nodePos && pos <= nodePos + node.nodeSize) {
        const start = nodePos + 1;
        result = { paragraph: paragraphIndex, run: 0, offset: Math.max(0, pos - start) };
        return false;
      }
    }
    return true;
  });
  return result ?? { paragraph: 0, run: 0, offset: 0 };
}

/** Selection range in document coordinates. Always returns start ≤ end. */
export function pmSelectionToRange(state: EditorState): DocxRange {
  const { from, to } = state.selection;
  const lo = Math.min(from, to);
  const hi = Math.max(from, to);
  return {
    start: pmPositionToDocx(state, lo),
    end: pmPositionToDocx(state, hi),
  };
}

const TEXT_MARK_NAMES = new Set(["bold", "italic", "underline", "strikethrough"]);

/**
 * Compute the set of inline marks active "at the selection" — the union
 * of marks on every text node in the range, falling back to stored
 * marks when the selection is collapsed. Used by the toolbar to show
 * pressed-state on Bold/Italic/Underline/Strike buttons.
 */
export function activeMarks(state: EditorState): Set<string> {
  const out = new Set<string>();
  const { from, to, empty, $from } = state.selection;
  if (empty) {
    const stored = state.storedMarks ?? $from.marks();
    for (const m of stored) {
      if (TEXT_MARK_NAMES.has(m.type.name)) out.add(m.type.name);
    }
    return out;
  }
  state.doc.nodesBetween(from, to, (node) => {
    if (!node.isText) return true;
    for (const m of node.marks) {
      if (TEXT_MARK_NAMES.has(m.type.name)) out.add(m.type.name);
    }
    return true;
  });
  return out;
}

/** Style id of a paragraph in the snapshot. Returns "Normal" if missing. */
export function paragraphStyle(snapshot: DocxSnapshot, paragraphIndex: number): string {
  const block = snapshot.root.body[paragraphIndex];
  if (!block || block.kind !== "paragraph") return "Normal";
  return block.properties.styleId ?? "Normal";
}

/** Paragraph index of the start of the current PM selection. */
export function currentParagraphIndex(state: EditorState): number {
  const { from } = state.selection;
  return pmPositionToDocx(state, from).paragraph;
}

export const PARAGRAPH_STYLES: ReadonlyArray<{ value: string; label: string }> = [
  { value: "Normal", label: "Normal" },
  { value: "Title", label: "Title" },
  { value: "Heading1", label: "Heading 1" },
  { value: "Heading2", label: "Heading 2" },
  { value: "Heading3", label: "Heading 3" },
];

/** Half-points multiplied by 2 — i.e. raw OOXML w:sz values. */
export const FONT_SIZES: ReadonlyArray<number> = [16, 18, 20, 22, 24, 28, 32, 36, 40, 48, 56, 72];

/** Common color picker palette. Hex strings without the `#`. */
export const COLOR_PALETTE: ReadonlyArray<{ name: string; hex: string }> = [
  { name: "Default", hex: "000000" },
  { name: "Gray", hex: "595959" },
  { name: "Red", hex: "C00000" },
  { name: "Orange", hex: "ED7D31" },
  { name: "Yellow", hex: "FFC000" },
  { name: "Green", hex: "70AD47" },
  { name: "Blue", hex: "2E75B6" },
  { name: "Purple", hex: "7030A0" },
];

/** OOXML w:highlight enum values (the only ones Word recognises). */
export const HIGHLIGHT_PALETTE: ReadonlyArray<{ name: string; label: string; swatch: string }> = [
  { name: "yellow", label: "Yellow", swatch: "#fff59d" },
  { name: "green", label: "Green", swatch: "#a5d6a7" },
  { name: "cyan", label: "Cyan", swatch: "#80deea" },
  { name: "magenta", label: "Magenta", swatch: "#f48fb1" },
  { name: "red", label: "Red", swatch: "#ef9a9a" },
  { name: "darkYellow", label: "Olive", swatch: "#827717" },
  { name: "lightGray", label: "Light gray", swatch: "#e0e0e0" },
];

/** Comment thread: top-level comment plus chronologically-ordered replies. */
export interface CommentThread {
  parent: DocxComment;
  replies: DocxComment[];
}

/**
 * Group comments into threads. Top-level comments (no `parentId`) become
 * thread parents; comments with a `parentId` are appended to their
 * parent's reply list in document order. Orphaned replies (parentId
 * pointing to a non-existent comment) are surfaced as their own
 * top-level threads so they don't silently disappear.
 */
export function commentThreads(snapshot: DocxSnapshot): CommentThread[] {
  const byId = new Map<string, DocxComment>();
  for (const c of snapshot.root.comments) byId.set(c.id, c);
  const threads = new Map<string, CommentThread>();
  const orphans: CommentThread[] = [];
  for (const c of snapshot.root.comments) {
    if (!c.parentId) {
      threads.set(c.id, { parent: c, replies: [] });
    }
  }
  for (const c of snapshot.root.comments) {
    if (!c.parentId) continue;
    const parent = byId.get(c.parentId);
    if (parent && threads.has(parent.id)) {
      threads.get(parent.id)!.replies.push(c);
    } else {
      orphans.push({ parent: c, replies: [] });
    }
  }
  return [...threads.values(), ...orphans];
}

/** Extract plain text of a comment's body (single-paragraph case primarily). */
export function commentPlainText(comment: DocxComment): string {
  let out = "";
  for (const block of comment.body) {
    if (block.kind !== "paragraph") continue;
    if (out.length > 0) out += "\n";
    out += paragraphText(block);
  }
  return out;
}

function paragraphText(p: Paragraph): string {
  let out = "";
  for (const child of p.children) {
    if (child.kind === "run") out += runText(child);
  }
  return out;
}

function runText(r: Run): string {
  let out = "";
  for (const c of r.children) if (c.kind === "text") out += c.text;
  return out;
}

/** Locate the paragraph that anchors a given comment id (via `comment-range-start`). */
export function commentParagraphIndex(snapshot: DocxSnapshot, commentId: string): number | null {
  const body = snapshot.root.body;
  for (let i = 0; i < body.length; i++) {
    const block = body[i];
    if (block.kind !== "paragraph") continue;
    if (paragraphContainsCommentStart(block, commentId)) return i;
  }
  return null;
}

function paragraphContainsCommentStart(p: Paragraph, commentId: string): boolean {
  for (const child of p.children) {
    if (child.kind === "comment-range-start" && child.commentId === commentId) return true;
  }
  return false;
}

/** Walk the snapshot and return every revision wrapper in document order. */
export function collectRevisions(snapshot: DocxSnapshot): RevisionWrapper[] {
  const out: RevisionWrapper[] = [];
  for (const block of snapshot.root.body) {
    if (block.kind !== "paragraph") continue;
    walkInlines(block.children, out);
  }
  return out;
}

function walkInlines(children: ReadonlyArray<InlineNode>, out: RevisionWrapper[]): void {
  for (const child of children) {
    if (child.kind === "revision") {
      out.push(child);
      walkInlines(child.children, out);
    }
  }
}

/** "Welcome to officeAI" → "Welcome to officeAI" (trim + collapse whitespace). */
export function snippet(text: string, max = 64): string {
  const t = text.replace(/\s+/g, " ").trim();
  if (t.length <= max) return t;
  return t.slice(0, max - 1) + "…";
}
