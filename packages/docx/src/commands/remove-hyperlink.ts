import { CommandError, type CommandHandler } from "@officeai/core";
import type {
  BlockNode,
  DocxDocument,
  DocxSnapshot,
  Hyperlink,
  InlineNode,
  Paragraph,
  Relationship,
  Table,
  TableCell,
  TableRow,
} from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { withoutRaw } from "./insert-table.js";
import type { RemoveHyperlinkPayload } from "./payloads.js";

const DOC_RELS_KEY = "word/document.xml";

/**
 * Unwrap a hyperlink: replace the `<w:hyperlink>` node with its
 * children (the runs spread back inline). When the unwrapped link
 * carried an external `relationshipId` AND no other body hyperlink
 * references the same id, the rel is removed from
 * `word/_rels/document.xml.rels` and `dirty.relationships` is set.
 *
 * The "is the rel still referenced" scan is body-only — header /
 * footer parts have their own rels parts in the W4 typed model and
 * aren't covered by this round (their rel-graph mutation is a
 * separate workstream). The W11 build-log entry documents the gap.
 */
export const removeHyperlinkHandler: CommandHandler<RemoveHyperlinkPayload, DocxSnapshot> = {
  type: "docx:remove-hyperlink",
  apply(snapshot, payload) {
    const { hyperlinkId } = payload;
    if (!hyperlinkId) {
      throw new CommandError("invalid-payload", "hyperlinkId is required");
    }

    const located = locateHyperlink(snapshot.root, hyperlinkId);
    if (!located) {
      throw new CommandError("unknown-target", `no hyperlink with id "${hyperlinkId}"`);
    }

    const { hyperlink, paragraph, replaceParagraph, path } = located;
    const newChildren: InlineNode[] = [];
    for (const inline of paragraph.children) {
      if (inline.kind === "hyperlink" && inline.id === hyperlinkId) {
        newChildren.push(...inline.children);
      } else {
        newChildren.push(inline);
      }
    }
    const newParagraph: Paragraph = { ...paragraph, children: newChildren };
    let nextDoc: DocxDocument = replaceParagraph(newParagraph);

    const relsToCheck = hyperlink.relationshipId;
    let relsDirtyChanged = false;
    if (relsToCheck) {
      const stillReferenced = countHyperlinkRelReferences(nextDoc, relsToCheck) > 0;
      if (!stillReferenced) {
        const docRels = nextDoc.relationships.get(DOC_RELS_KEY) ?? [];
        const filtered = docRels.filter((r) => r.id !== relsToCheck);
        if (filtered.length !== docRels.length) {
          const newMap = new Map<string, ReadonlyArray<Relationship>>(nextDoc.relationships);
          newMap.set(DOC_RELS_KEY, filtered);
          nextDoc = { ...nextDoc, relationships: newMap };
          relsDirtyChanged = true;
        }
      }
    }

    const nextDirty = relsDirtyChanged
      ? { body: true, relationships: withAddition(snapshot.dirty.relationships, DOC_RELS_KEY) }
      : { body: true };
    const next = evolveSnapshot(snapshot, nextDoc, nextDirty);
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: paragraph.id,
        path,
        field: "children",
        summary: `−hyperlink (${relsToCheck ? `rel=${relsToCheck}${relsDirtyChanged ? " removed" : " kept"}` : "anchor"})`,
      }),
    };
  },
};

function withAddition(prev: ReadonlySet<string>, member: string): ReadonlySet<string> {
  const next = new Set(prev);
  next.add(member);
  return next;
}

interface LocatedHyperlink {
  readonly hyperlink: Hyperlink;
  readonly paragraph: Paragraph;
  readonly path: ReadonlyArray<string | number>;
  readonly replaceParagraph: (next: Paragraph) => DocxDocument;
}

function locateHyperlink(doc: DocxDocument, hyperlinkId: string): LocatedHyperlink | null {
  for (let i = 0; i < doc.body.length; i++) {
    const block = doc.body[i];
    if (block.kind === "paragraph") {
      const hl = findHyperlinkInParagraph(block, hyperlinkId);
      if (hl) {
        return {
          hyperlink: hl,
          paragraph: block,
          path: ["body", i],
          replaceParagraph: (next) => {
            const body = doc.body.slice();
            body[i] = next;
            return { ...doc, body };
          },
        };
      }
    } else if (block.kind === "table") {
      const found = findHyperlinkInTable(block, hyperlinkId, ["body", i]);
      if (found) {
        const replaceParagraph = (next: Paragraph): DocxDocument => {
          const newTable = found.rebuild(next);
          const body = doc.body.slice();
          body[i] = newTable;
          return { ...doc, body };
        };
        return {
          hyperlink: found.hyperlink,
          paragraph: found.paragraph,
          path: found.path,
          replaceParagraph,
        };
      }
    }
  }
  return null;
}

function findHyperlinkInParagraph(p: Paragraph, hyperlinkId: string): Hyperlink | null {
  for (const inline of p.children) {
    if (inline.kind === "hyperlink" && inline.id === hyperlinkId) return inline;
  }
  return null;
}

interface InnerLocatedHl {
  readonly hyperlink: Hyperlink;
  readonly paragraph: Paragraph;
  readonly path: ReadonlyArray<string | number>;
  readonly rebuild: (next: Paragraph) => Table;
}

function findHyperlinkInTable(
  table: Table,
  hyperlinkId: string,
  basePath: ReadonlyArray<string | number>
): InnerLocatedHl | null {
  for (let r = 0; r < table.rows.length; r++) {
    const row = table.rows[r];
    for (let c = 0; c < row.cells.length; c++) {
      const cell = row.cells[c];
      for (let b = 0; b < cell.body.length; b++) {
        const block = cell.body[b];
        const path = [...basePath, "rows", r, "cells", c, "body", b];
        if (block.kind === "paragraph") {
          const hl = findHyperlinkInParagraph(block, hyperlinkId);
          if (hl) {
            const rebuild = (next: Paragraph): Table => rewriteCellBlock(table, r, c, b, next);
            return { hyperlink: hl, paragraph: block, path, rebuild };
          }
        } else if (block.kind === "table") {
          const inner = findHyperlinkInTable(block, hyperlinkId, path);
          if (inner) {
            const rebuild = (next: Paragraph): Table => {
              const newInner = inner.rebuild(next);
              return rewriteCellBlock(table, r, c, b, newInner);
            };
            return { hyperlink: inner.hyperlink, paragraph: inner.paragraph, path: inner.path, rebuild };
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

function countHyperlinkRelReferences(doc: DocxDocument, relId: string): number {
  let n = 0;
  for (const block of doc.body) {
    n += countInBlock(block, relId);
  }
  return n;
}

function countInBlock(block: BlockNode, relId: string): number {
  if (block.kind === "paragraph") {
    let n = 0;
    for (const inline of block.children) {
      if (inline.kind === "hyperlink" && inline.relationshipId === relId) n++;
    }
    return n;
  }
  if (block.kind === "table") {
    let n = 0;
    for (const row of block.rows) {
      for (const cell of row.cells) {
        for (const b of cell.body) n += countInBlock(b, relId);
      }
    }
    return n;
  }
  return 0;
}
