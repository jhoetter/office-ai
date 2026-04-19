import type { Mark, MarkType, Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import type { Step } from "prosemirror-transform";
import { ReplaceStep, ReplaceAroundStep, AddMarkStep, RemoveMarkStep } from "prosemirror-transform";
import type { Command, CommandLite } from "@officeai/core";
import type { DocxPosition, DocxSelection } from "../model/types.js";
import type { TextFormat } from "../commands/payloads.js";

/**
 * Translate a ProseMirror transaction into a (possibly empty) list of
 * commands for the bus. We intentionally support a small subset that
 * covers normal typing/editing; unsupported steps are reported via the
 * `unsupported` channel so the EditorView can show a toast.
 *
 * Coordinate model: PM positions are global integer offsets across the
 * document. We translate them to `DocxPosition`s via the previous PM
 * doc's structure (paragraph index = nth top-level paragraph; offset =
 * character offset within the paragraph's text).
 */

export interface UnsupportedTx {
  reason: string;
  step?: Step;
}

export interface TranslationResult {
  commands: Array<Command | CommandLite>;
  unsupported: UnsupportedTx[];
}

export interface TranslationOptions {
  agentId?: string;
  source?: "human" | "agent" | "system";
  /**
   * Edit mode for the transaction:
   *   - `"edit"` (default) — insertions and deletions become plain
   *     `docx:insert-text` / `docx:delete-range` commands; this is
   *     the historical behaviour.
   *   - `"suggest"` — insertions become `docx:insert-text-tracked`
   *     and single-paragraph deletions become
   *     `docx:delete-range-tracked` (`<w:ins>` / `<w:del>` revision
   *     wrappers); the changes show up in the tracked-changes UI.
   *     `author` MUST be supplied; multi-paragraph deletions and
   *     paste-with-block-content fall back to the unsupported
   *     channel because they need additional revision plumbing
   *     (paragraph-mark deletes, multi-block tracked inserts).
   */
  mode?: "edit" | "suggest";
  /** Author attribution for tracked changes; required when `mode === "suggest"`. */
  author?: string;
}

export function transactionToCommands(
  tx: Transaction,
  before: EditorState,
  opts: TranslationOptions = {}
): TranslationResult {
  const out: TranslationResult = { commands: [], unsupported: [] };
  const source = opts.source ?? "human";
  const agentId = opts.agentId;
  const mode = opts.mode ?? "edit";
  const author = opts.author ?? "";
  let docCursor = before.doc;

  for (const step of tx.steps) {
    const handled = handleStep(step, docCursor, before, out, source, agentId, mode, author);
    if (!handled) {
      out.unsupported.push({ reason: "unsupported step", step });
    }
    const result = step.apply(docCursor);
    if (result.failed === null && result.doc) {
      docCursor = result.doc;
    }
  }

  return out;
}

function handleStep(
  step: Step,
  doc: PMNode,
  before: EditorState,
  out: TranslationResult,
  source: "human" | "agent" | "system",
  agentId: string | undefined,
  mode: "edit" | "suggest",
  author: string
): boolean {
  if (step instanceof ReplaceStep) {
    return handleReplaceStep(step, doc, before, out, source, agentId, mode, author);
  }
  if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
    return handleMarkStep(step, doc, out, source, agentId);
  }
  if (step instanceof ReplaceAroundStep) {
    out.unsupported.push({ reason: "structural replace around (lists/blockquote) deferred", step });
    return false;
  }
  return false;
}

function handleReplaceStep(
  step: ReplaceStep,
  doc: PMNode,
  before: EditorState,
  out: TranslationResult,
  source: "human" | "agent" | "system",
  agentId: string | undefined,
  mode: "edit" | "suggest",
  author: string
): boolean {
  const { from, to } = step;
  const sliceContent = step.slice.content;
  const startPos = positionFromPM(doc, from);
  const endPos = positionFromPM(doc, to);

  if (!startPos || !endPos) {
    out.unsupported.push({ reason: "could not translate PM positions to model positions", step });
    return false;
  }

  if (sliceHasUnsupportedBlock(sliceContent)) {
    out.unsupported.push({ reason: "slice contains a non-paragraph block (table / opaque)", step });
    return false;
  }

  const blockSegments = sliceBlockSegments(sliceContent);
  const insertedHasBlocks = blockSegments.length > 1;
  const insertedText = blockSegments.length > 0 ? blockSegments[0] : "";

  // Suggesting mode: route insert/delete through the tracked
  // variants so they materialise as `<w:ins>` / `<w:del>` revision
  // wrappers. Multi-block paste and multi-paragraph deletes are
  // out of MVP scope (they need extra revision plumbing — paragraph
  // mark deletes, multi-block tracked inserts) and fall through
  // to the unsupported channel so the toolbar can warn the user.
  if (mode === "suggest") {
    if (!author) {
      out.unsupported.push({
        reason: "suggesting mode requires a non-empty author; falling back to direct edit",
        step,
      });
      return false;
    }
    if (insertedHasBlocks) {
      out.unsupported.push({
        reason: "tracked multi-block paste is not yet supported in Suggesting mode",
        step,
      });
      return false;
    }
    if (from === to) {
      if (insertedText.length === 0) return true;
      out.commands.push(
        wrap(
          {
            type: "docx:insert-text-tracked",
            payload: { at: startPos, text: insertedText, author },
          },
          source,
          agentId
        )
      );
      return true;
    }
    if (startPos.paragraph !== endPos.paragraph) {
      out.unsupported.push({
        reason: "tracked deletes across paragraph boundaries are not yet supported in Suggesting mode",
        step,
      });
      return false;
    }
    out.commands.push(
      wrap(
        {
          type: "docx:delete-range-tracked",
          payload: { range: makeRange(startPos, endPos), author },
        },
        source,
        agentId
      )
    );
    if (insertedText.length > 0) {
      out.commands.push(
        wrap(
          {
            type: "docx:insert-text-tracked",
            payload: { at: startPos, text: insertedText, author },
          },
          source,
          agentId
        )
      );
    }
    return true;
  }

  // Insertion at a cursor (no range to delete).
  if (from === to) {
    if (insertedHasBlocks) {
      emitMultiBlockPaste(blockSegments, startPos, out, source, agentId);
      return true;
    }
    if (insertedText.length === 0) return true;
    emitInsertWithMarkReassertion(insertedText, startPos, step, doc, before, from, out, source, agentId);
    return true;
  }

  // Range deletion + nothing to insert.
  if (blockSegments.length === 0 || (blockSegments.length === 1 && insertedText.length === 0)) {
    out.commands.push(
      wrap({ type: "docx:delete-range", payload: { range: makeRange(startPos, endPos) } }, source, agentId)
    );
    return true;
  }

  // Replace = delete + insert.
  out.commands.push(
    wrap({ type: "docx:delete-range", payload: { range: makeRange(startPos, endPos) } }, source, agentId)
  );
  if (insertedHasBlocks) {
    emitMultiBlockPaste(blockSegments, startPos, out, source, agentId);
  } else {
    emitInsertWithMarkReassertion(insertedText, startPos, step, doc, before, from, out, source, agentId);
  }
  return true;
}

/**
 * Emit `insert-text` and, when warranted, a follow-up `format-range` so
 * the inserted text inherits the ambient marks. PM's `tr.insertText`
 * automatically applies `storedMarks` to the slice's text node, so
 * reading marks off the slice is the most reliable signal; we fall back
 * to "marks at the insertion point" for programmatic insertions that
 * arrive without slice-level marks.
 */
function emitInsertWithMarkReassertion(
  text: string,
  startPos: DocxPosition,
  step: ReplaceStep,
  doc: PMNode,
  before: EditorState,
  from: number,
  out: TranslationResult,
  source: "human" | "agent" | "system",
  agentId: string | undefined
): void {
  out.commands.push(wrap({ type: "docx:insert-text", payload: { at: startPos, text } }, source, agentId));
  const format = ambientFormatAtInsert(step, doc, before, from);
  if (format && !isEmptyFormat(format)) {
    const endParagraphOffset = (startPos.offset ?? 0) + text.length;
    const range: DocxSelection = {
      start: { paragraph: startPos.paragraph, offset: startPos.offset ?? 0 },
      end: { paragraph: startPos.paragraph, offset: endParagraphOffset },
    };
    out.commands.push(wrap({ type: "docx:format-range", payload: { range, format } }, source, agentId));
  }
}

/**
 * Translate a multi-block paste slice into a sequence of `insert-text`
 * + `insert-paragraph` commands. The slice is laid out as N paragraph
 * segments; we emit:
 *   1. insert-text for segment[0] at the cursor,
 *   2. insert-paragraph at (cursor + segment[0].length) — splits the paragraph,
 *   3. for every middle segment: insert-text at the start of the new
 *      paragraph, then insert-paragraph at end of that text,
 *   4. insert-text for the last segment at the start of the final paragraph.
 *
 * We compute positions in the model coordinate space, not PM offsets,
 * so the result is independent of PM's `+2 per paragraph` accounting.
 */
function emitMultiBlockPaste(
  segments: string[],
  startPos: DocxPosition,
  out: TranslationResult,
  source: "human" | "agent" | "system",
  agentId: string | undefined
): void {
  if (segments.length < 2) return;
  let paragraphIndex = startPos.paragraph;
  let offset = startPos.offset ?? 0;

  // First segment goes into the existing paragraph at the cursor.
  const first = segments[0];
  if (first.length > 0) {
    out.commands.push(
      wrap(
        { type: "docx:insert-text", payload: { at: { paragraph: paragraphIndex, offset }, text: first } },
        source,
        agentId
      )
    );
    offset += first.length;
  }
  // Split the current paragraph at the boundary, creating paragraph N+1.
  out.commands.push(
    wrap(
      { type: "docx:insert-paragraph", payload: { at: { paragraph: paragraphIndex, offset } } },
      source,
      agentId
    )
  );
  paragraphIndex++;
  offset = 0;

  // Middle segments: insert text into the freshly-created paragraph,
  // then split again to start the next.
  for (let i = 1; i < segments.length - 1; i++) {
    const seg = segments[i];
    if (seg.length > 0) {
      out.commands.push(
        wrap(
          { type: "docx:insert-text", payload: { at: { paragraph: paragraphIndex, offset }, text: seg } },
          source,
          agentId
        )
      );
    }
    out.commands.push(
      wrap(
        {
          type: "docx:insert-paragraph",
          payload: { at: { paragraph: paragraphIndex, offset: seg.length } },
        },
        source,
        agentId
      )
    );
    paragraphIndex++;
    offset = 0;
  }

  // Last segment slots into the final new paragraph.
  const last = segments[segments.length - 1];
  if (last.length > 0) {
    out.commands.push(
      wrap(
        { type: "docx:insert-text", payload: { at: { paragraph: paragraphIndex, offset }, text: last } },
        source,
        agentId
      )
    );
  }
}

function handleMarkStep(
  step: AddMarkStep | RemoveMarkStep,
  doc: PMNode,
  out: TranslationResult,
  source: "human" | "agent" | "system",
  agentId: string | undefined
): boolean {
  const { from, to } = step;
  const startPos = positionFromPM(doc, from);
  const endPos = positionFromPM(doc, to);
  if (!startPos || !endPos) {
    out.unsupported.push({ reason: "could not translate PM mark positions to model", step });
    return false;
  }
  const value = step instanceof AddMarkStep;
  const format = formatFromMark(step.mark, value);
  if (!format) {
    out.unsupported.push({ reason: `unsupported mark type: ${step.mark.type.name}`, step });
    return false;
  }
  out.commands.push(
    wrap(
      { type: "docx:format-range", payload: { range: makeRange(startPos, endPos), format } },
      source,
      agentId
    )
  );
  return true;
}

function wrap(
  cmd: { type: string; payload: unknown },
  source: "human" | "agent" | "system",
  agentId: string | undefined
): Command {
  const base = { type: cmd.type, payload: cmd.payload, source };
  return agentId ? ({ ...base, agentId } as Command) : (base as Command);
}

function makeRange(start: DocxPosition, end: DocxPosition): DocxSelection {
  return { start, end };
}

function formatFromMark(mark: Mark, value: boolean): TextFormat | null {
  const t: MarkType = mark.type;
  switch (t.name) {
    case "bold":
      return { bold: value };
    case "italic":
      return { italic: value };
    case "underline":
      return { underline: value };
    case "strikethrough":
      return { strike: value };
    case "color":
      return value ? { color: String(mark.attrs.rgb ?? "") } : { color: "" };
    case "highlight":
      return value ? { highlight: String(mark.attrs.name ?? "") } : { highlight: "" };
    case "font_family":
      return value ? { fontFamily: String(mark.attrs.family ?? "") } : { fontFamily: "" };
    case "font_size":
      return value ? { fontSize: Number(mark.attrs.halfPoints ?? 0) } : {};
    default:
      return null;
  }
}

/**
 * Compute a `TextFormat` representing the marks the inserted text
 * should carry. Order of preference:
 *   1. marks on the inserted slice's first text node — PM applies
 *      `storedMarks` here, so this captures continued-typing inside a
 *      formatted span.
 *   2. `EditorState.storedMarks` on the previous state.
 *   3. marks at the insertion point (`$pos.marks()`).
 */
function ambientFormatAtInsert(
  step: ReplaceStep,
  doc: PMNode,
  before: EditorState,
  from: number
): TextFormat | null {
  const sliceMarks = firstTextMarks(step.slice.content);
  if (sliceMarks && sliceMarks.length > 0) {
    return formatFromMarks(sliceMarks);
  }
  const stored = before.storedMarks ?? null;
  if (stored && stored.length > 0) {
    return formatFromMarks(stored);
  }
  const ambient = marksAtPos(doc, from);
  if (ambient.length > 0) return formatFromMarks(ambient);
  return null;
}

function firstTextMarks(content: import("prosemirror-model").Fragment): readonly Mark[] | null {
  let marks: readonly Mark[] | null = null;
  content.descendants((node) => {
    if (marks !== null) return false;
    if (node.isText) {
      marks = node.marks;
      return false;
    }
    return true;
  });
  return marks;
}

function marksAtPos(doc: PMNode, pos: number): readonly Mark[] {
  try {
    return doc.resolve(pos).marks();
  } catch {
    return [];
  }
}

/**
 * Reduce a list of PM marks into a single `TextFormat`. Structural
 * marks (hyperlink, comment, revision) are paragraph / wrapper concepts
 * in our model, not run properties — they're filtered out so the
 * follow-up `format-range` only carries actual typographic toggles.
 */
function formatFromMarks(marks: readonly Mark[]): TextFormat {
  const out: TextFormat = {};
  for (const m of marks) {
    const f = formatFromMark(m, true);
    if (!f) continue;
    Object.assign(out, f);
  }
  return out;
}

function isEmptyFormat(f: TextFormat): boolean {
  return (
    f.bold === undefined &&
    f.italic === undefined &&
    f.underline === undefined &&
    f.strike === undefined &&
    f.fontFamily === undefined &&
    f.fontSize === undefined &&
    f.color === undefined &&
    f.highlight === undefined
  );
}

/**
 * Walk a slice and return the plain text of every top-level paragraph
 * block, in order. A slice with one paragraph collapses to a one-element
 * array; a slice that bare-text would otherwise be inline-only collapses
 * to `[text]` so callers see a single segment.
 */
function sliceBlockSegments(content: import("prosemirror-model").Fragment): string[] {
  const segments: string[] = [];
  let inlineText = "";
  let sawBlock = false;
  for (let i = 0; i < content.childCount; i++) {
    const child = content.child(i);
    if (child.isBlock && child.type.name === "paragraph") {
      sawBlock = true;
      segments.push(extractPlainText(child.content));
      continue;
    }
    if (child.isInline || child.isText) {
      inlineText += extractInlineText(child);
    }
  }
  if (!sawBlock) {
    return inlineText.length > 0 ? [inlineText] : [];
  }
  return segments;
}

function sliceHasUnsupportedBlock(content: import("prosemirror-model").Fragment): boolean {
  let bad = false;
  content.forEach((child) => {
    if (bad) return;
    if (child.isBlock && child.type.name !== "paragraph") {
      bad = true;
    }
  });
  return bad;
}

/**
 * Translate a global PM offset to a DocxPosition. The PM doc is a
 * sequence of top-level paragraph (or atom) blocks; for paragraphs we
 * compute a 0-based character offset into the paragraph's plain text.
 */
function positionFromPM(doc: PMNode, pos: number): DocxPosition | null {
  let paragraphIndex = -1;
  let result: DocxPosition | null = null;
  doc.descendants((node, nodePos) => {
    if (node.type.name === "doc") return true;
    if (result) return false;
    if (node.isBlock && node.type.name === "paragraph") {
      paragraphIndex++;
      const start = nodePos + 1;
      const end = nodePos + node.nodeSize - 1;
      if (pos >= nodePos && pos <= nodePos + node.nodeSize) {
        const offset = Math.max(0, Math.min(end, pos) - start);
        // We deliberately omit `run` so handlers treat `offset` as a
        // global paragraph-text offset. The PM doc collapses runs +
        // revision wrappers + other inline children into a single
        // text stream (revision_mark is a Mark, not a Node), so we
        // can't honestly point at a specific run from here. The
        // insert / delete handlers walk paragraph children to find
        // the right splice point. (Returning `run: 0` instead would
        // make every tracked insertion land in run 0 and the bus
        // would reverse character order — see Suggesting-mode
        // regression test.)
        result = { paragraph: paragraphIndex, offset };
        return false;
      }
      return false;
    }
    if (node.isBlock) {
      paragraphIndex++;
      if (pos >= nodePos && pos <= nodePos + node.nodeSize) {
        result = { paragraph: paragraphIndex, offset: 0 };
        return false;
      }
      return false;
    }
    return true;
  });
  if (result) return result;
  return { paragraph: Math.max(0, paragraphIndex), offset: 0 };
}

function extractPlainText(content: import("prosemirror-model").Fragment): string {
  let s = "";
  content.descendants((node) => {
    if (node.isText) {
      s += node.text ?? "";
      return false;
    }
    if (node.type.name === "tab") {
      s += "\t";
      return false;
    }
    if (node.type.name === "hard_break") {
      return false;
    }
    return true;
  });
  return s;
}

function extractInlineText(node: PMNode): string {
  if (node.isText) return node.text ?? "";
  if (node.type.name === "tab") return "\t";
  return "";
}
