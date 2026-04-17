import { CommandError, type CommandHandler, type DiffChange } from "@officeai/core";
import type {
  BlockNode,
  DocxDocument,
  DocxPosition,
  DocxSnapshot,
  InlineNode,
  Paragraph,
  Run,
  RunChild,
  TextLeaf,
} from "../model/types.js";
import {
  buildDiff,
  buildDiffMulti,
  emptyRun,
  evolveSnapshot,
  removeBlocks,
  replaceBlock,
} from "./helpers.js";
import type { DeleteRangePayload } from "./payloads.js";

/**
 * `docx:delete-range` handler.
 *
 * Single-paragraph ranges splice and trim the affected runs in place.
 * Multi-paragraph ranges (`start.paragraph !== end.paragraph`) trim the
 * start paragraph from the start boundary to its end, drop every
 * intermediate block (paragraphs and non-paragraph blocks alike), trim
 * the end paragraph from its beginning to the end boundary, and merge
 * the trimmed start + trimmed end into a single paragraph that
 * preserves the start paragraph's `id` and `properties`.
 */
export const deleteRangeHandler: CommandHandler<DeleteRangePayload, DocxSnapshot> = {
  type: "docx:delete-range",
  apply(snapshot, payload, ctx) {
    const { range } = payload;
    const [startPos, endPos] = orderPositions(range.start, range.end);
    if (startPos.paragraph === endPos.paragraph) {
      return applySingleParagraph(snapshot, startPos, endPos);
    }
    return applyMultiParagraph(snapshot, startPos, endPos, ctx.mintNodeId);
  },
};

function orderPositions(a: DocxPosition, b: DocxPosition): [DocxPosition, DocxPosition] {
  if (a.paragraph < b.paragraph) return [a, b];
  if (a.paragraph > b.paragraph) return [b, a];
  return [a, b];
}

function applySingleParagraph(snapshot: DocxSnapshot, startPos: DocxPosition, endPos: DocxPosition) {
  const idx = startPos.paragraph;
  if (idx < 0 || idx >= snapshot.root.body.length) {
    throw new CommandError("invalid-position", `paragraph index ${idx} out of range`);
  }
  const block = snapshot.root.body[idx];
  if (block.kind !== "paragraph") {
    throw new CommandError("not-paragraph", `block at ${idx} is not a paragraph`);
  }
  const startOff = paragraphTextOffset(block, startPos.run, startPos.offset ?? 0);
  const endOff = paragraphTextOffset(block, endPos.run, endPos.offset ?? 0);
  const lo = Math.min(startOff, endOff);
  const hi = Math.max(startOff, endOff);
  if (lo === hi) {
    return {
      next: { ...snapshot, revision: snapshot.revision + 1 },
      diff: buildDiff(snapshot.revision, snapshot.revision + 1, {
        kind: "node-updated",
        nodeId: block.id,
        path: ["body", idx],
        field: "noop",
        summary: "no-op delete-range",
      }),
    };
  }
  const updated = deleteWithinParagraph(block, lo, hi);
  const nextDoc = replaceBlock(snapshot.root, idx, updated);
  const next = evolveSnapshot(snapshot, nextDoc, { body: true });
  return {
    next,
    diff: buildDiff(snapshot.revision, next.revision, {
      kind: "node-updated",
      nodeId: block.id,
      path: ["body", idx],
      field: "text",
      summary: `−${hi - lo} chars`,
    }),
  };
}

function applyMultiParagraph(
  snapshot: DocxSnapshot,
  startPos: DocxPosition,
  endPos: DocxPosition,
  mintNodeId: () => string
) {
  const startIdx = startPos.paragraph;
  const endIdx = endPos.paragraph;
  if (startIdx < 0 || endIdx >= snapshot.root.body.length) {
    throw new CommandError("invalid-position", `paragraph indices ${startIdx}..${endIdx} out of range`);
  }
  const startBlock = snapshot.root.body[startIdx];
  const endBlock = snapshot.root.body[endIdx];
  if (startBlock.kind !== "paragraph") {
    throw new CommandError("not-paragraph", `block at ${startIdx} is not a paragraph`);
  }
  if (endBlock.kind !== "paragraph") {
    throw new CommandError("not-paragraph", `block at ${endIdx} is not a paragraph`);
  }

  const startTextOff = paragraphTextOffset(startBlock, startPos.run, startPos.offset ?? 0);
  const endTextOff = paragraphTextOffset(endBlock, endPos.run, endPos.offset ?? 0);

  const startTotal = paragraphTextLength(startBlock);
  const trimmedStart = deleteWithinParagraph(startBlock, startTextOff, startTotal);
  const trimmedEnd = deleteWithinParagraph(endBlock, 0, endTextOff);

  const merged = mergeParagraphs(trimmedStart, trimmedEnd, mintNodeId);

  const droppedBlocks: BlockNode[] = [];
  for (let i = startIdx + 1; i <= endIdx; i++) {
    droppedBlocks.push(snapshot.root.body[i]);
  }

  let doc: DocxDocument = snapshot.root;
  doc = replaceBlock(doc, startIdx, merged);
  doc = removeBlocks(doc, startIdx + 1, endIdx + 1);

  const next = evolveSnapshot(snapshot, doc, { body: true });

  const changes: DiffChange[] = [];
  changes.push({
    kind: "node-updated",
    nodeId: merged.id,
    path: ["body", startIdx],
    field: "text",
    summary: `merged with paragraph ${endIdx}`,
  });
  for (let i = 0; i < droppedBlocks.length; i++) {
    const dropped = droppedBlocks[i];
    changes.push({
      kind: "node-deleted",
      nodeId: dropped.id,
      path: ["body", startIdx + 1 + i],
      summary: `deleted ${dropped.kind}`,
    });
  }

  return {
    next,
    diff: buildDiffMulti(snapshot.revision, next.revision, changes),
  };
}

function mergeParagraphs(start: Paragraph, end: Paragraph, mintNodeId: () => string): Paragraph {
  const children: InlineNode[] = [...start.children, ...end.children];
  const hasRun = children.some((c) => c.kind === "run");
  if (!hasRun) {
    children.push(emptyRun(mintNodeId, {}));
  }
  return {
    kind: "paragraph",
    id: start.id,
    properties: start.properties,
    children,
  };
}

/**
 * Convert a `(runIndex, localOffset)` pair into a paragraph-wide character
 * offset. When `runIndex` is undefined we treat `localOffset` as already
 * paragraph-wide and clamp it to the paragraph's text length.
 */
function paragraphTextOffset(p: Paragraph, runIndex: number | undefined, localOffset: number): number {
  if (runIndex === undefined) {
    return Math.max(0, Math.min(localOffset, paragraphTextLength(p)));
  }
  let offset = 0;
  let i = 0;
  for (const child of p.children) {
    if (child.kind !== "run") {
      i++;
      continue;
    }
    if (i === runIndex) {
      let consumed = 0;
      for (const c of child.children) {
        if (c.kind !== "text") continue;
        if (localOffset <= consumed + c.text.length) {
          return offset + (localOffset - consumed);
        }
        offset += c.text.length;
        consumed += c.text.length;
      }
      return offset + Math.max(0, localOffset - consumed);
    }
    for (const c of child.children) {
      if (c.kind === "text") offset += c.text.length;
    }
    i++;
  }
  return offset;
}

function paragraphTextLength(p: Paragraph): number {
  let n = 0;
  for (const child of p.children) {
    if (child.kind !== "run") continue;
    for (const c of child.children) {
      if (c.kind === "text") n += c.text.length;
    }
  }
  return n;
}

function deleteWithinParagraph(p: Paragraph, lo: number, hi: number): Paragraph {
  let cursor = 0;
  const finalChildren = p.children.map((c) => {
    if (c.kind !== "run") return c;
    const before = cursor;
    const after = before + runTextLength(c);
    if (hi <= before || lo >= after) {
      cursor = after;
      return c;
    }
    const localLo = Math.max(0, lo - before);
    const localHi = Math.min(after - before, hi - before);
    const [updated] = deleteWithinRun(c, localLo, localHi);
    cursor = after;
    return updated;
  });
  return { ...p, children: finalChildren };
}

function runTextLength(r: Run): number {
  let n = 0;
  for (const c of r.children) if (c.kind === "text") n += c.text.length;
  return n;
}

function deleteWithinRun(r: Run, lo: number, hi: number): [Run, number] {
  let consumed = 0;
  const out: RunChild[] = [];
  for (const child of r.children) {
    if (child.kind !== "text") {
      out.push(child);
      continue;
    }
    const len = child.text.length;
    const start = consumed;
    const end = consumed + len;
    if (hi <= start || lo >= end) {
      out.push(child);
    } else {
      const lLo = Math.max(0, lo - start);
      const lHi = Math.min(len, hi - start);
      const before = child.text.slice(0, lLo);
      const after = child.text.slice(lHi);
      const merged = before + after;
      const updated: TextLeaf = { ...child, text: merged };
      out.push(updated);
    }
    consumed = end;
  }
  return [{ ...r, children: out }, hi - lo];
}
