import { CommandError, type CommandHandler } from "@officeai/core";
import type {
  BlockNode,
  DocxPosition,
  DocxSnapshot,
  InlineNode,
  Paragraph,
  RevisionWrapper,
  Run,
  RunChild,
  TextLeaf,
} from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceBlock } from "./helpers.js";
import type { DeleteRangeTrackedPayload } from "./payloads.js";

/**
 * Tracked-changes ("Suggesting" mode) variant of `docx:delete-range`.
 *
 * Instead of removing the targeted text from the paragraph, every
 * text leaf inside the range is wrapped in a `<w:del>` revision
 * (its leaves flipped to `isDelText: true` so the serializer emits
 * `<w:delText>`). Visually the deleted text shows up as a
 * strike-through suggestion in Word / Google Docs / our own
 * renderer. `accept-change` later removes it for real;
 * `reject-change` unwraps it back to plain text.
 *
 * MVP scope: single-paragraph ranges only. Multi-paragraph deletes
 * are rejected with a `not-implemented` error because they require
 * tagging the intermediate paragraph marks (`<w:rPr><w:del/></w:rPr>`
 * on the paragraph's `<w:pPr>`) so Word knows to merge the
 * paragraphs on accept; that surface is deferred to a follow-up.
 */
export const deleteRangeTrackedHandler: CommandHandler<DeleteRangeTrackedPayload, DocxSnapshot> = {
  type: "docx:delete-range-tracked",
  apply(snapshot, payload, ctx) {
    const { range, author } = payload;
    if (!author) {
      throw new CommandError("invalid-payload", "delete-range-tracked requires a non-empty author");
    }
    const [startPos, endPos] = orderPositions(range.start, range.end);
    if (startPos.paragraph !== endPos.paragraph) {
      throw new CommandError(
        "not-implemented",
        "tracked deletes across paragraph boundaries are not yet supported; split the range or accept previous changes first"
      );
    }
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
          summary: "no-op delete-range-tracked",
        }),
      };
    }

    const date = payload.date ?? new Date().toISOString();
    const revisionId = payload.revisionId ?? mintRevisionId(snapshot);

    const updated = wrapRangeAsDeletion(block, lo, hi, author, date, revisionId, ctx.mintNodeId);
    const nextDoc = replaceBlock(snapshot.root, idx, updated);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: block.id,
        path: ["body", idx],
        field: "tracked-delete",
        summary: `−del ${hi - lo} chars`,
      }),
    };
  },
};

function orderPositions(a: DocxPosition, b: DocxPosition): [DocxPosition, DocxPosition] {
  if (a.paragraph < b.paragraph) return [a, b];
  if (a.paragraph > b.paragraph) return [b, a];
  return [a, b];
}

function mintRevisionId(snapshot: DocxSnapshot): string {
  const used = new Set<number>();
  for (const b of snapshot.root.body) visitBlockForRevisions(b, used);
  for (const part of snapshot.root.headersAndFooters) {
    for (const b of part.body) visitBlockForRevisions(b, used);
  }
  let id = 1;
  while (used.has(id)) id++;
  return String(id);
}

function visitBlockForRevisions(block: BlockNode, used: Set<number>): void {
  if (block.kind !== "paragraph") return;
  for (const child of block.children) collectFromInline(child, used);
}

function collectFromInline(node: InlineNode, used: Set<number>): void {
  if (node.kind === "revision") {
    const n = Number(node.revisionId);
    if (Number.isFinite(n) && n > 0 && Math.floor(n) === n) used.add(n);
    for (const c of node.children) collectFromInline(c, used);
  }
}

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
  for (const child of p.children) n += inlineTextLength(child);
  return n;
}

function inlineTextLength(node: InlineNode): number {
  if (node.kind === "run") {
    let n = 0;
    for (const c of node.children) if (c.kind === "text") n += c.text.length;
    return n;
  }
  if (node.kind === "revision" || node.kind === "hyperlink") {
    let n = 0;
    for (const c of node.children) n += inlineTextLength(c);
    return n;
  }
  return 0;
}

/**
 * Walk every run in document order and split each one at the range
 * boundaries (`lo`, `hi`). Run segments that fall *inside* the range
 * are emitted as fresh runs whose text leaves are flipped to
 * `isDelText: true`, then collected into a single `<w:del>` revision
 * wrapper inserted at the boundary of the first deleted segment.
 *
 * We deliberately keep one wrapper covering the whole contiguous
 * range (rather than one per run-segment): Word writes deletions
 * the same way and the accept-change handler expects each
 * `revisionId` to address a single wrapper.
 */
function wrapRangeAsDeletion(
  p: Paragraph,
  lo: number,
  hi: number,
  author: string,
  date: string,
  revisionId: string,
  mintNodeId: () => string
): Paragraph {
  let cursor = 0;
  const out: InlineNode[] = [];
  const deletedRuns: Run[] = [];
  let wrapperInserted = false;

  for (const child of p.children) {
    if (child.kind !== "run") {
      out.push(child);
      continue;
    }
    const start = cursor;
    const end = start + runTextLength(child);
    if (hi <= start || lo >= end) {
      // Run lies entirely outside the deletion range.
      out.push(child);
      cursor = end;
      continue;
    }

    const localLo = Math.max(0, lo - start);
    const localHi = Math.min(end - start, hi - start);
    const segments = splitRunForDeletion(child, localLo, localHi, mintNodeId);

    for (const seg of segments) {
      if (seg.kind === "before" || seg.kind === "after") {
        if (seg.run.children.length > 0) out.push(seg.run);
      } else {
        deletedRuns.push(seg.run);
        if (!wrapperInserted) {
          // Reserve the wrapper's slot at the position of the first
          // deleted segment. We push a marker now and replace it
          // with the assembled wrapper after the loop so we can
          // gather all deleted run segments first.
          out.push(WRAPPER_MARKER);
          wrapperInserted = true;
        }
      }
    }
    cursor = end;
  }

  if (deletedRuns.length === 0) {
    // Range targeted only non-text content; nothing to wrap.
    return p;
  }

  const wrapper: RevisionWrapper = {
    kind: "revision",
    id: mintNodeId(),
    revisionType: "del",
    author,
    date,
    revisionId,
    children: deletedRuns,
  };
  const finalChildren = out.map((n) => (n === WRAPPER_MARKER ? wrapper : n));
  return { ...p, children: finalChildren };
}

/**
 * Sentinel inline node used as a placeholder for the eventual
 * `<w:del>` wrapper. Replaced with the real wrapper after we've
 * walked every run and collected all deleted segments. Picking a
 * sentinel rather than building the wrapper upfront keeps the run
 * walk single-pass and lets us handle ranges that span multiple
 * runs without juggling indices.
 */
const WRAPPER_MARKER = { kind: "__wrapper_marker__" } as unknown as InlineNode;

type DeletionSegment =
  | { kind: "before"; run: Run }
  | { kind: "deleted"; run: Run }
  | { kind: "after"; run: Run };

/**
 * Split a single run into up to three segments around the
 * `[localLo, localHi)` deletion window. The middle segment carries
 * the soon-to-be-wrapped text with `isDelText: true`; the
 * surrounding segments preserve formatting and the original run's
 * non-text leaves. Empty segments are returned and filtered by the
 * caller so the run list stays compact.
 */
function splitRunForDeletion(
  run: Run,
  localLo: number,
  localHi: number,
  mintNodeId: () => string
): DeletionSegment[] {
  let consumed = 0;
  const beforeChildren: RunChild[] = [];
  const deletedChildren: RunChild[] = [];
  const afterChildren: RunChild[] = [];

  for (const child of run.children) {
    if (child.kind !== "text") {
      // Non-text run children bind to the segment containing their
      // textual position. Since they're zero-width in the offset
      // model, push to whichever segment matches the cursor.
      if (consumed < localLo) beforeChildren.push(child);
      else if (consumed >= localHi) afterChildren.push(child);
      else deletedChildren.push(child);
      continue;
    }
    const len = child.text.length;
    const start = consumed;
    const end = consumed + len;

    const beforePart = child.text.slice(0, Math.max(0, Math.min(len, localLo - start)));
    const middlePart = child.text.slice(
      Math.max(0, Math.min(len, localLo - start)),
      Math.max(0, Math.min(len, localHi - start))
    );
    const afterPart = child.text.slice(Math.max(0, Math.min(len, localHi - start)));

    if (beforePart.length > 0) {
      beforeChildren.push({ ...child, text: beforePart, isDelText: false });
    }
    if (middlePart.length > 0) {
      const delLeaf: TextLeaf = {
        ...child,
        id: mintNodeId(),
        text: middlePart,
        isDelText: true,
        // <w:delText> always needs xml:space="preserve" semantics
        // because Word expects the strike-through to retain its
        // exact whitespace contribution.
        xmlSpacePreserve: true,
      };
      deletedChildren.push(delLeaf);
    }
    if (afterPart.length > 0) {
      afterChildren.push({ ...child, id: mintNodeId(), text: afterPart, isDelText: false });
    }
    consumed = end;
    void start;
    void end;
  }

  const segments: DeletionSegment[] = [];
  segments.push({ kind: "before", run: { ...run, children: beforeChildren } });
  if (deletedChildren.length > 0) {
    segments.push({
      kind: "deleted",
      run: { ...run, id: mintNodeId(), children: deletedChildren },
    });
  }
  segments.push({ kind: "after", run: { ...run, id: mintNodeId(), children: afterChildren } });
  return segments;
}

function runTextLength(r: Run): number {
  let n = 0;
  for (const c of r.children) if (c.kind === "text") n += c.text.length;
  return n;
}
