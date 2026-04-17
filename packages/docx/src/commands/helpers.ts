import type {
  BlockNode,
  DocxDirtyFlags,
  DocxDocument,
  DocxSnapshot,
  Paragraph,
  Run,
  RunProperties,
  TextLeaf,
} from "../model/types.js";
import type { DocumentDiff } from "@officeai/core";

/**
 * Build a new snapshot from `prev` with the given mutator. Increments
 * `revision`. Caller passes which dirty flags to set.
 */
export function evolveSnapshot(
  prev: DocxSnapshot,
  next: DocxDocument,
  dirty: Partial<DocxDirtyFlags>
): DocxSnapshot {
  return {
    format: "docx",
    revision: prev.revision + 1,
    root: next,
    partHashes: prev.partHashes,
    container: prev.container,
    dirty: { ...prev.dirty, ...dirty },
  };
}

/** Replace one block at `index` in body. */
export function replaceBlock(doc: DocxDocument, index: number, block: BlockNode): DocxDocument {
  const body = doc.body.slice();
  body[index] = block;
  return { ...doc, body };
}

/** Insert a block at `index`. */
export function insertBlock(doc: DocxDocument, index: number, block: BlockNode): DocxDocument {
  const body = doc.body.slice();
  body.splice(index, 0, block);
  return { ...doc, body };
}

/** Remove a range of blocks (inclusive start, exclusive end). */
export function removeBlocks(doc: DocxDocument, start: number, end: number): DocxDocument {
  const body = doc.body.slice();
  body.splice(start, end - start);
  return { ...doc, body };
}

/** Update one paragraph at `index`, asserting it is a paragraph. */
export function withParagraph(
  doc: DocxDocument,
  index: number,
  fn: (p: Paragraph) => Paragraph
): DocxDocument {
  const block = doc.body[index];
  if (!block || block.kind !== "paragraph") {
    throw new Error(`block at index ${index} is not a paragraph`);
  }
  return replaceBlock(doc, index, fn(block));
}

/**
 * Iterate inline runs of a paragraph in document order. Returns a flat list
 * with reference indexes useful for updating.
 */
export function getRuns(p: Paragraph): { runIndex: number; run: Run }[] {
  const out: { runIndex: number; run: Run }[] = [];
  for (let i = 0; i < p.children.length; i++) {
    const c = p.children[i];
    if (c.kind === "run") out.push({ runIndex: i, run: c });
  }
  return out;
}

/** Return a fresh empty Run with optional properties. */
export function emptyRun(mintNodeId: () => string, properties: RunProperties = {}): Run {
  return {
    kind: "run",
    id: mintNodeId(),
    properties,
    children: [],
  };
}

/** Return a TextLeaf. Sets xml:space="preserve" only when the text has leading/trailing whitespace. */
export function textLeaf(mintNodeId: () => string, text: string): TextLeaf {
  return {
    kind: "text",
    id: mintNodeId(),
    text,
    xmlSpacePreserve: /^\s|\s$/.test(text) || text.length === 0,
  };
}

/** Concatenate all text leaves of a paragraph into a single plain string. */
export function paragraphPlainText(p: Paragraph): string {
  let out = "";
  for (const c of p.children) {
    if (c.kind === "run") {
      for (const child of c.children) {
        if (child.kind === "text") out += child.text;
      }
    } else if (c.kind === "hyperlink") {
      for (const r of c.children) {
        for (const child of r.children) {
          if (child.kind === "text") out += child.text;
        }
      }
    } else if (c.kind === "revision") {
      for (const ic of c.children) {
        if (ic.kind === "run") {
          for (const child of ic.children) {
            if (child.kind === "text") out += child.text;
          }
        }
      }
    }
  }
  return out;
}

/** Build a single-change diff. */
export function buildDiff(
  prevRevision: number,
  nextRevision: number,
  change: DocumentDiff["changes"][number]
): DocumentDiff {
  return {
    format: "docx",
    fromRevision: prevRevision,
    toRevision: nextRevision,
    changes: [change],
  };
}
