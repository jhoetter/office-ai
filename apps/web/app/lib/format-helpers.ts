import type { EditorState } from "prosemirror-state";
import {
  resolveEffectiveRpr,
  type DocxComment,
  type DocxSnapshot,
  type RevisionWrapper,
  type InlineNode,
  type Paragraph,
  type Run,
  type RunProperties,
} from "@officeai/docx";

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

/** Sentinel returned by `activeMarkAttr` when the selection spans runs
 * with conflicting values for a given mark attribute (e.g. half of the
 * range is 11pt and half is 14pt). The toolbar dropdowns render this
 * as "Mixed" / "—" so the user knows the selection is heterogeneous. */
export const MIXED = Symbol("mixed");
export type MaybeMixed<T> = T | typeof MIXED | undefined;

/**
 * Compute the dominant value of `markName.attrName` across the current
 * PM selection. Returns:
 *   - `undefined` when no run in the selection carries the mark,
 *   - `MIXED` when runs carry the mark but with conflicting values,
 *   - the value otherwise.
 *
 * Collapsed selections inspect the stored marks (the marks that would
 * apply to the next typed character), matching the semantics of
 * `activeMarks`.
 */
export function activeMarkAttr<T = unknown>(
  state: EditorState,
  markName: string,
  attrName: string
): MaybeMixed<T> {
  const { from, to, empty, $from } = state.selection;
  if (empty) {
    const stored = state.storedMarks ?? $from.marks();
    for (const m of stored) {
      if (m.type.name === markName) return m.attrs[attrName] as T;
    }
    return undefined;
  }
  let seen: T | undefined;
  let mixed = false;
  let sawMark = false;
  state.doc.nodesBetween(from, to, (node) => {
    if (mixed) return false;
    if (!node.isText) return true;
    const mark = node.marks.find((m) => m.type.name === markName);
    if (!mark) return true;
    sawMark = true;
    const value = mark.attrs[attrName] as T;
    if (seen === undefined) {
      seen = value;
    } else if (seen !== value) {
      mixed = true;
    }
    return true;
  });
  if (mixed) return MIXED;
  if (!sawMark) return undefined;
  return seen;
}

/**
 * P3.1 / W3 — selection-aware run attribute lookup that falls back to
 * the typed style cascade when no direct PM mark carries the attribute.
 *
 * Word inherits run formatting from a four-level cascade
 * (`docDefaults.rPrDefault` → `paragraph.styleId` chain → paragraph's
 * own `pPr.rPr` → run's `rPr`). The PM mark check above only sees the
 * leaf; the rest lives in `word/styles.xml`. When a paragraph carries
 * `styleId="Heading1"` and the run inside has empty `<w:rPr/>`, the
 * resolver returns the inherited `Calibri 16pt` and the toolbar shows
 * `16` / `Calibri` instead of the placeholder.
 *
 * Behaviour:
 *   - Direct PM mark found → identical to `activeMarkAttr`.
 *   - No direct mark → resolve effective rPr per paragraph in the
 *     selection, project to `inheritedSelector(rPr)`, return MIXED if
 *     paragraphs disagree.
 *   - Empty selection collapses to "the paragraph the caret is in".
 */
export function activeRunAttr<T>(
  state: EditorState,
  markName: string,
  attrName: string,
  snapshot: DocxSnapshot | null,
  inheritedSelector: (rPr: RunProperties) => T | undefined
): MaybeMixed<T> {
  const direct = activeMarkAttr<T>(state, markName, attrName);
  if (direct !== undefined) return direct;
  if (!snapshot) return undefined;

  const paragraphIndices = collectSelectionParagraphIndices(state);
  if (paragraphIndices.length === 0) return undefined;

  let seen: T | undefined;
  let mixed = false;
  for (const idx of paragraphIndices) {
    const resolved = resolveEffectiveRpr(snapshot, idx, 0);
    const value = inheritedSelector(resolved);
    if (value === undefined) continue;
    if (seen === undefined) seen = value;
    else if (seen !== value) {
      mixed = true;
      break;
    }
  }
  if (mixed) return MIXED;
  return seen;
}

function collectSelectionParagraphIndices(state: EditorState): number[] {
  const { from, to } = state.selection;
  const indices: number[] = [];
  let paragraphIndex = -1;
  state.doc.descendants((node, nodePos) => {
    if (node.type.name !== "paragraph") return true;
    paragraphIndex++;
    const nodeEnd = nodePos + node.nodeSize;
    if (nodeEnd >= from && nodePos <= to) indices.push(paragraphIndex);
    return false;
  });
  return indices;
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

/**
 * Derive the paragraph-style picker contents from the **document** rather
 * than the hard-coded list. We don't have a typed `styles.xml` model
 * yet (planned for a later phase), so we approximate by walking the
 * loaded body and collecting every `styleId` actually used in the
 * doc. This means a freshly-uploaded thesis with `berschrift1`-`6` and
 * `Untertitel` actually shows those entries instead of the demo
 * defaults — which is the user-visible bug from the Masterthesis
 * screenshot.
 *
 * Sort order: Title-family first, Heading 1-9 (English) next, then
 * German Word equivalents, then Normal, then everything else
 * alphabetically. The current `activeStyle` is always included even
 * when it doesn't appear in the body (e.g. an empty document whose
 * caret sits on a Heading1 stub).
 */
export function paragraphStyleOptions(
  snapshot: DocxSnapshot | null,
  activeStyle: string
): ReadonlyArray<{ value: string; label: string }> {
  const used = new Set<string>();
  if (snapshot) {
    for (const block of snapshot.root.body) {
      if (block.kind !== "paragraph") continue;
      const id = block.properties.styleId;
      if (id) used.add(id);
    }
  }
  // Always include the curated defaults so a brand-new doc still has
  // sensible options.
  for (const s of PARAGRAPH_STYLES) used.add(s.value);
  if (activeStyle) used.add(activeStyle);

  const seen = Array.from(used);
  return seen
    .map((value) => ({ value, label: humanizeStyleId(value) }))
    .sort((a, b) => styleSortKey(a.value) - styleSortKey(b.value) || a.label.localeCompare(b.label));
}

function styleSortKey(id: string): number {
  if (id === "Title" || id === "Titel") return 0;
  if (id === "Subtitle" || id === "Untertitel") return 1;
  const headingMatch = /^(Heading|berschrift)(\d)$/.exec(id);
  if (headingMatch) return 10 + Number(headingMatch[2]);
  if (id === "Normal") return 50;
  if (id === "TOCHeading" || id === "Verzeichnis") return 60;
  if (/^TOC\d/.test(id)) return 61;
  return 100;
}

function humanizeStyleId(id: string): string {
  // Word's German style ids drop the leading 'Ü' from `Überschrift` and
  // are otherwise camel-case. Surface them in their canonical English
  // form when possible so the dropdown reads naturally even when the
  // doc was authored in Word DE.
  switch (id) {
    case "Title":
    case "Titel":
      return "Title";
    case "Subtitle":
    case "Untertitel":
      return "Subtitle";
    case "Normal":
      return "Normal";
    case "TOCHeading":
    case "Verzeichnis":
      return "TOC heading";
  }
  const heading = /^(Heading|berschrift)(\d)$/.exec(id);
  if (heading) return `Heading ${heading[2]}`;
  const toc = /^TOC(\d)$/.exec(id);
  if (toc) return `TOC ${toc[1]}`;
  // Generic camel-case → "Camel case" formatter as a fallback.
  const spaced = id.replace(/([a-z])([A-Z])/g, "$1 $2").replace(/_/g, " ");
  return spaced.charAt(0).toUpperCase() + spaced.slice(1);
}

/** Half-points multiplied by 2 — i.e. raw OOXML w:sz values. */
export const FONT_SIZES: ReadonlyArray<number> = [16, 18, 20, 22, 24, 28, 32, 36, 40, 48, 56, 72];

/** Default font family list shown in the toolbar when the document
 *  doesn't carry an explicit fonts table or its fonts can't be parsed. */
export const FONT_FAMILIES: ReadonlyArray<string> = [
  "Calibri",
  "Cambria",
  "Times New Roman",
  "Arial",
  "Helvetica",
  "Georgia",
  "Verdana",
  "Tahoma",
  "Courier New",
  "Consolas",
];

/** Discover an existing numbering definition for a given list kind. We
 * don't yet auto-mint a `<w:num>`/`<w:abstractNum>` pair when the doc
 * has no numbering.xml at all (P1.4 deferred this), so the toolbar
 * falls back to "unsupported" in that case rather than silently
 * no-op'ing. */
export function discoverNumId(
  snapshot: DocxSnapshot | null,
  kind: "bullet" | "ordered"
): { numId: number; ilvl: number } | null {
  const numbering = snapshot?.root.numbering;
  if (!numbering) return null;
  for (const num of numbering.nums.values()) {
    const abstract = numbering.abstractNums.get(num.abstractNumId);
    if (!abstract) continue;
    const level0 = abstract.levels.find((l) => l.ilvl === 0) ?? abstract.levels[0];
    if (!level0) continue;
    const fmt = level0.numFmt ?? "decimal";
    const isBullet = fmt === "bullet";
    if (kind === "bullet" && isBullet) return { numId: num.numId, ilvl: 0 };
    if (kind === "ordered" && !isBullet) return { numId: num.numId, ilvl: 0 };
  }
  return null;
}

/**
 * Look up the stable `paragraphId` of the paragraph containing the
 * caret. Returns `null` when the caret sits inside an opaque/atomic
 * block (table, image, opaque carrier).
 */
export function currentParagraphId(state: EditorState): string | null {
  const { from } = state.selection;
  let paragraphId: string | null = null;
  state.doc.descendants((node, nodePos) => {
    if (paragraphId !== null) return false;
    if (node.type.name === "paragraph") {
      if (from >= nodePos && from <= nodePos + node.nodeSize) {
        const id = node.attrs.paragraphId;
        if (typeof id === "string" && id.length > 0) paragraphId = id;
        return false;
      }
    }
    return true;
  });
  return paragraphId;
}

/** Alignment of the paragraph containing the caret, or `null` when the
 * caret is not inside a paragraph (or the paragraph carries no
 * explicit alignment). */
export function currentParagraphAlignment(
  state: EditorState
): "left" | "center" | "right" | "justify" | null {
  const { from } = state.selection;
  let alignment: "left" | "center" | "right" | "justify" | null = null;
  state.doc.descendants((node, nodePos) => {
    if (node.type.name === "paragraph") {
      if (from >= nodePos && from <= nodePos + node.nodeSize) {
        const v = node.attrs.alignment;
        if (v === "left" || v === "center" || v === "right" || v === "justify") {
          alignment = v;
        }
        return false;
      }
    }
    return true;
  });
  return alignment;
}

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

/**
 * Revision wrapper paired with the plain-text it carries.
 *
 * The Word-style margin balloons render as `<author> hat gelöscht: <text>`
 * (or `eingefügt:` for insertions), so the UI needs the inline text the
 * wrapper covers. Computing it once here keeps the React component free
 * of model traversal.
 */
export interface RevisionWithPreview extends RevisionWrapper {
  readonly previewText: string;
}

export function collectRevisionsWithPreview(snapshot: DocxSnapshot): RevisionWithPreview[] {
  return collectRevisions(snapshot).map((rev) => ({
    ...rev,
    previewText: revisionPlainText(rev),
  }));
}

function revisionPlainText(rev: RevisionWrapper): string {
  let acc = "";
  const visit = (children: ReadonlyArray<InlineNode>): void => {
    for (const child of children) {
      switch (child.kind) {
        case "run":
          for (const leaf of child.children) {
            if (leaf.kind === "text") acc += leaf.text;
          }
          break;
        case "revision":
          visit(child.children);
          break;
        case "hyperlink":
          for (const inner of child.children) {
            if (inner.kind === "run") {
              for (const leaf of inner.children) {
                if (leaf.kind === "text") acc += leaf.text;
              }
            }
          }
          break;
        case "comment-range-start":
        case "comment-range-end":
        case "comment-reference":
        case "opaque-inline":
          break;
        default: {
          const _exhaustive: never = child;
          void _exhaustive;
        }
      }
    }
  };
  visit(rev.children);
  return acc;
}

/** "Welcome to officeAI" → "Welcome to officeAI" (trim + collapse whitespace). */
export function snippet(text: string, max = 64): string {
  const t = text.replace(/\s+/g, " ").trim();
  if (t.length <= max) return t;
  return t.slice(0, max - 1) + "…";
}
