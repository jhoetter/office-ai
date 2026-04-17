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
}

export function transactionToCommands(
  tx: Transaction,
  before: EditorState,
  opts: TranslationOptions = {}
): TranslationResult {
  const out: TranslationResult = { commands: [], unsupported: [] };
  const source = opts.source ?? "human";
  const agentId = opts.agentId;
  let docCursor = before.doc;

  for (const step of tx.steps) {
    const handled = handleStep(step, docCursor, out, source, agentId);
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
  out: TranslationResult,
  source: "human" | "agent" | "system",
  agentId: string | undefined
): boolean {
  if (step instanceof ReplaceStep) {
    return handleReplaceStep(step, doc, out, source, agentId);
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
  out: TranslationResult,
  source: "human" | "agent" | "system",
  agentId: string | undefined
): boolean {
  const { from, to } = step;
  const sliceContent = step.slice.content;
  const insertedText = extractPlainText(sliceContent);
  const insertedHasBlocks = sliceHasBlocks(sliceContent);
  const startPos = positionFromPM(doc, from);
  const endPos = positionFromPM(doc, to);

  if (!startPos || !endPos) {
    out.unsupported.push({ reason: "could not translate PM positions to model positions", step });
    return false;
  }

  if (from === to) {
    if (insertedHasBlocks) {
      // Treat any block-bearing slice insertion at a cursor as a paragraph split (Enter key).
      out.commands.push(
        wrap(
          { type: "docx:insert-paragraph", payload: { at: startPos } },
          source,
          agentId
        )
      );
      return true;
    }
    if (insertedText.length === 0) return true;
    out.commands.push(
      wrap(
        { type: "docx:insert-text", payload: { at: startPos, text: insertedText } },
        source,
        agentId
      )
    );
    return true;
  }

  if (insertedText.length === 0 && !insertedHasBlocks) {
    out.commands.push(
      wrap(
        { type: "docx:delete-range", payload: { range: makeRange(startPos, endPos) } },
        source,
        agentId
      )
    );
    return true;
  }

  // Replace = delete + insert.
  out.commands.push(
    wrap({ type: "docx:delete-range", payload: { range: makeRange(startPos, endPos) } }, source, agentId)
  );
  if (insertedText.length > 0) {
    out.commands.push(
      wrap(
        { type: "docx:insert-text", payload: { at: startPos, text: insertedText } },
        source,
        agentId
      )
    );
  }
  if (insertedHasBlocks) {
    out.commands.push(
      wrap({ type: "docx:insert-paragraph", payload: { at: startPos } }, source, agentId)
    );
  }
  return true;
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
      const start = nodePos + 1; // inside the paragraph
      const end = nodePos + node.nodeSize - 1;
      if (pos >= nodePos && pos <= nodePos + node.nodeSize) {
        const offset = Math.max(0, Math.min(end, pos) - start);
        result = { paragraph: paragraphIndex, run: 0, offset };
        return false;
      }
      return false;
    }
    if (node.isBlock) {
      paragraphIndex++;
      if (pos >= nodePos && pos <= nodePos + node.nodeSize) {
        result = { paragraph: paragraphIndex, run: 0, offset: 0 };
        return false;
      }
      return false;
    }
    return true;
  });
  if (result) return result;
  // Position past the last block.
  return { paragraph: Math.max(0, paragraphIndex), run: 0, offset: 0 };
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
      // hard_break inside a paragraph is not the same as "Enter"; skip for now.
      return false;
    }
    return true;
  });
  return s;
}

function sliceHasBlocks(content: import("prosemirror-model").Fragment): boolean {
  let found = false;
  content.descendants((node) => {
    if (node.isBlock && node.type.name === "paragraph") {
      found = true;
      return false;
    }
    return !found;
  });
  return found;
}
