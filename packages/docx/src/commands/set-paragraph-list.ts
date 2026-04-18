import { CommandError, type CommandHandler } from "@officeai/core";
import type {
  BlockNode,
  DocxDocument,
  DocxSnapshot,
  Paragraph,
  ParagraphProperties,
  Table,
  TableCell,
  TableRow,
} from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { withoutRaw } from "./insert-table.js";
import type { SetParagraphListPayload } from "./payloads.js";

/**
 * Set (or replace) a paragraph's list reference. The paragraph may live
 * directly in the body or inside any table cell (recursively). The
 * change is paragraph-local: only `dirty.body` is set. The numbering
 * **definitions** part (`word/numbering.xml`) is NOT mutated — the
 * command merely points the paragraph at an existing `<w:num>`
 * instance, which is exactly how Word's "apply list" toolbar works.
 *
 * If the document carries no `word/numbering.xml`, the command rejects
 * with `unknown-target`. Auto-creating the part (would need a fresh
 * `<w:abstractNum>` + `<w:num>` pair plus `[Content_Types].xml`
 * registration) is deferred — see the W10 build-log entry.
 */
export const setParagraphListHandler: CommandHandler<SetParagraphListPayload, DocxSnapshot> = {
  type: "docx:set-paragraph-list",
  apply(snapshot, payload) {
    const { paragraphId, numId, ilvl } = payload;
    if (!paragraphId) {
      throw new CommandError("invalid-payload", "paragraphId is required");
    }
    if (!Number.isInteger(numId) || numId <= 0) {
      throw new CommandError("invalid-payload", `numId must be a positive integer (got ${numId})`);
    }
    if (!Number.isInteger(ilvl) || ilvl < 0) {
      throw new CommandError("invalid-payload", `ilvl must be a non-negative integer (got ${ilvl})`);
    }

    const numbering = snapshot.root.numbering;
    if (!numbering) {
      throw new CommandError(
        "unknown-target",
        `document has no word/numbering.xml; cannot apply list (numId=${numId}, ilvl=${ilvl})`
      );
    }
    const numInstance = numbering.nums.get(numId);
    if (!numInstance) {
      throw new CommandError("unknown-target", `no <w:num> with numId=${numId} in word/numbering.xml`);
    }
    const abstractNum = numbering.abstractNums.get(numInstance.abstractNumId);
    if (abstractNum && abstractNum.levels.length > 0 && ilvl >= abstractNum.levels.length) {
      throw new CommandError(
        "invalid-payload",
        `ilvl ${ilvl} exceeds abstractNum levels (count=${abstractNum.levels.length})`
      );
    }

    const located = locateParagraph(snapshot.root, paragraphId);
    if (!located) {
      throw new CommandError("unknown-target", `no paragraph with id "${paragraphId}"`);
    }

    const updatedProps = applyNumberingToProps(located.paragraph.properties, numId, ilvl);
    const updatedParagraph: Paragraph = { ...located.paragraph, properties: updatedProps };
    const nextDoc = located.replace(updatedParagraph);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: located.paragraph.id,
        path: located.path,
        field: "numbering",
        summary: `set list numId=${numId} ilvl=${ilvl}`,
      }),
    };
  },
};

/**
 * Replace the paragraph's typed `numbering` field and strip any opaque
 * `<w:numPr>` carrier left over from parse. The serializer keys off
 * the typed field when no opaque carrier is present (see
 * `serializeParagraphProperties`), which is exactly the contract we
 * want here.
 */
function applyNumberingToProps(props: ParagraphProperties, numId: number, ilvl: number): ParagraphProperties {
  const opaqueProps = props.opaqueProps?.filter((o) => o.tag !== "w:numPr");
  const next: ParagraphProperties = {
    ...props,
    numbering: { numId, ilvl },
    ...(opaqueProps && opaqueProps.length > 0 ? { opaqueProps } : { opaqueProps: undefined }),
  };
  // `undefined` opaqueProps must be removed from the result so structural
  // equality on round-trip stays clean.
  if (next.opaqueProps === undefined) {
    const { opaqueProps: _drop, ...rest } = next;
    void _drop;
    return rest;
  }
  return next;
}

/**
 * Recursively locate a paragraph by id anywhere in the body — including
 * inside nested table cells. Returns a closure that rebuilds every
 * enclosing structure (cells, rows, tables) with `raw` cleared on
 * touched tables, so the change reaches the serializer at every depth.
 */
export interface LocatedParagraph {
  readonly paragraph: Paragraph;
  readonly path: ReadonlyArray<string | number>;
  readonly replace: (next: Paragraph) => DocxDocument;
}

export function locateParagraph(doc: DocxDocument, paragraphId: string): LocatedParagraph | null {
  for (let i = 0; i < doc.body.length; i++) {
    const block = doc.body[i];
    if (block.kind === "paragraph" && block.id === paragraphId) {
      return {
        paragraph: block,
        path: ["body", i],
        replace: (next) => {
          const body = doc.body.slice();
          body[i] = next;
          return { ...doc, body };
        },
      };
    }
    if (block.kind === "table") {
      const found = findInTable(block, paragraphId, ["body", i]);
      if (found) {
        const replace = (next: Paragraph): DocxDocument => {
          const newTable = found.rebuild(next);
          const body = doc.body.slice();
          body[i] = newTable;
          return { ...doc, body };
        };
        return { paragraph: found.paragraph, path: found.path, replace };
      }
    }
  }
  return null;
}

interface InnerLocated {
  readonly paragraph: Paragraph;
  readonly path: ReadonlyArray<string | number>;
  readonly rebuild: (next: Paragraph) => Table;
}

function findInTable(
  table: Table,
  paragraphId: string,
  basePath: ReadonlyArray<string | number>
): InnerLocated | null {
  for (let r = 0; r < table.rows.length; r++) {
    const row = table.rows[r];
    for (let c = 0; c < row.cells.length; c++) {
      const cell = row.cells[c];
      for (let b = 0; b < cell.body.length; b++) {
        const block = cell.body[b];
        const path = [...basePath, "rows", r, "cells", c, "body", b];
        if (block.kind === "paragraph" && block.id === paragraphId) {
          const rebuild = (next: Paragraph): Table => rewriteCellBlock(table, r, c, b, next);
          return { paragraph: block, path, rebuild };
        }
        if (block.kind === "table") {
          const inner = findInTable(block, paragraphId, path);
          if (inner) {
            const rebuild = (nextPara: Paragraph): Table => {
              const newInner = inner.rebuild(nextPara);
              return rewriteCellBlock(table, r, c, b, newInner);
            };
            return { paragraph: inner.paragraph, path: inner.path, rebuild };
          }
        }
      }
    }
  }
  return null;
}

function rewriteCellBlock(table: Table, r: number, c: number, b: number, next: BlockNode): Table {
  const row = table.rows[r];
  const cell = row.cells[c];
  const newBody = cell.body.slice();
  newBody[b] = next;
  const newCell: TableCell = { ...cell, body: newBody };
  const newCells = row.cells.slice();
  newCells[c] = newCell;
  const newRow: TableRow = { ...row, cells: newCells };
  const newRows = table.rows.slice();
  newRows[r] = newRow;
  return withoutRaw({ ...table, rows: newRows });
}
