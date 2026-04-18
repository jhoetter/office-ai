import { CommandError, type CommandHandler, type IdMinter } from "@officeai/core";
import type {
  DocxSnapshot,
  InlineNode,
  PageBreakLeaf,
  Paragraph,
  Run,
  RunChild,
  TextLeaf,
} from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { locateParagraph } from "./set-paragraph-list.js";

/**
 * P3.5 / W18 — splice a typed `<w:br w:type="page"/>` into a body
 * paragraph at a flat-text offset.
 *
 * Mirrors {@link insertPageNumberHandler}'s splitting strategy: the
 * targeting run is split at the offset, a fresh single-leaf run is
 * spliced in, and the trailing half is appended after it. Run
 * properties carry over so the surrounding font / color survives the
 * split.
 *
 * The page chunker (`chunkIntoPages`) recognises the new leaf and
 * advances its page count on the next snapshot.
 */
export interface InsertPageBreakPayload {
  paragraphId: string;
  offset: number;
}

export const insertPageBreakHandler: CommandHandler<InsertPageBreakPayload, DocxSnapshot> = {
  type: "docx:insert-page-break",
  apply(snapshot, payload, ctx) {
    const { paragraphId, offset } = payload;
    if (!paragraphId) {
      throw new CommandError("invalid-payload", "paragraphId is required");
    }
    if (!Number.isInteger(offset) || offset < 0) {
      throw new CommandError("invalid-payload", `offset must be a non-negative integer (got ${offset})`);
    }

    const located = locateParagraph(snapshot.root, paragraphId);
    if (!located) {
      throw new CommandError(
        "unknown-target",
        `no body paragraph with id "${paragraphId}" — page breaks are only supported in the body in P3.5`
      );
    }

    const updated = insertBreakIntoParagraph(located.paragraph, offset, ctx.mintNodeId);
    const nextDoc = located.replace(updated);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: located.paragraph.id,
        path: located.path,
        field: "page-break",
        summary: `inserted page break at offset ${offset}`,
      }),
    };
  },
};

function insertBreakIntoParagraph(p: Paragraph, offset: number, mintNodeId: IdMinter): Paragraph {
  const breakLeaf: PageBreakLeaf = { kind: "page-break", id: mintNodeId() };
  const flatLength = paragraphFlatLength(p);
  const clamped = Math.max(0, Math.min(offset, flatLength));

  const newChildren: InlineNode[] = [];
  let consumed = 0;
  let placed = false;

  for (const child of p.children) {
    if (placed || child.kind !== "run") {
      newChildren.push(child);
      continue;
    }
    const runLength = runFlatLength(child);
    if (clamped >= consumed && clamped <= consumed + runLength) {
      const localOffset = clamped - consumed;
      const { before, after } = splitRunAt(child, localOffset, mintNodeId);
      const breakRun: Run = {
        kind: "run",
        id: mintNodeId(),
        properties: child.properties,
        children: [breakLeaf],
      };
      if (before) newChildren.push(before);
      newChildren.push(breakRun);
      if (after) newChildren.push(after);
      placed = true;
    } else {
      newChildren.push(child);
    }
    consumed += runLength;
  }

  if (!placed) {
    newChildren.push({ kind: "run", id: mintNodeId(), properties: {}, children: [breakLeaf] });
  }
  return { ...p, children: newChildren };
}

function paragraphFlatLength(p: Paragraph): number {
  let n = 0;
  for (const c of p.children) {
    if (c.kind === "run") n += runFlatLength(c);
  }
  return n;
}

function runFlatLength(r: Run): number {
  let n = 0;
  for (const c of r.children) {
    if (c.kind === "text") n += c.text.length;
  }
  return n;
}

interface SplitRun {
  readonly before: Run | null;
  readonly after: Run | null;
}

function splitRunAt(run: Run, localOffset: number, mintNodeId: IdMinter): SplitRun {
  const total = runFlatLength(run);
  if (localOffset <= 0) {
    return { before: null, after: run };
  }
  if (localOffset >= total) {
    return { before: run, after: null };
  }
  const beforeChildren: RunChild[] = [];
  const afterChildren: RunChild[] = [];
  let consumed = 0;
  let split = false;
  for (const child of run.children) {
    if (split) {
      afterChildren.push(child);
      continue;
    }
    if (child.kind !== "text") {
      beforeChildren.push(child);
      continue;
    }
    const len = child.text.length;
    if (localOffset >= consumed && localOffset <= consumed + len) {
      const localInLeaf = localOffset - consumed;
      const beforeText = child.text.slice(0, localInLeaf);
      const afterText = child.text.slice(localInLeaf);
      if (beforeText.length > 0) {
        const beforeLeaf: TextLeaf = { ...child, text: beforeText };
        beforeChildren.push(beforeLeaf);
      }
      if (afterText.length > 0) {
        const afterLeaf: TextLeaf = { ...child, id: mintNodeId(), text: afterText };
        afterChildren.push(afterLeaf);
      }
      split = true;
    } else {
      beforeChildren.push(child);
    }
    consumed += len;
  }
  const before: Run | null =
    beforeChildren.length > 0
      ? { kind: "run", id: run.id, properties: run.properties, children: beforeChildren }
      : null;
  const after: Run | null =
    afterChildren.length > 0
      ? { kind: "run", id: mintNodeId(), properties: run.properties, children: afterChildren }
      : null;
  return { before, after };
}
