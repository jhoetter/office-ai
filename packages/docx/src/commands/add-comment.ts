import { CommandError, type CommandHandler } from "@officeai/core";
import type {
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  DocxComment,
  DocxSnapshot,
  InlineNode,
  Paragraph,
  Run,
} from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceBlock, textLeaf } from "./helpers.js";
import type { AddCommentPayload } from "./payloads.js";

export const addCommentHandler: CommandHandler<AddCommentPayload, DocxSnapshot> = {
  type: "docx:add-comment",
  apply(snapshot, payload, ctx) {
    const { range, text, author, initials } = payload;
    if (range.start.paragraph !== range.end.paragraph) {
      throw new CommandError(
        "multi-paragraph-comment",
        "Multi-paragraph add-comment is P1 (this session: single-paragraph ranges only)."
      );
    }
    const idx = range.start.paragraph;
    const block = snapshot.root.body[idx];
    if (!block || block.kind !== "paragraph") {
      throw new CommandError("invalid-position", `paragraph index ${idx} is not a paragraph`);
    }

    const commentId = mintCommentId(snapshot.root.comments);
    const start: CommentRangeStart = { kind: "comment-range-start", id: ctx.mintNodeId(), commentId };
    const end: CommentRangeEnd = { kind: "comment-range-end", id: ctx.mintNodeId(), commentId };
    const ref: CommentReference = { kind: "comment-reference", id: ctx.mintNodeId(), commentId };

    const paragraphWithMarkers = wrapParagraphWithCommentMarkers(block, start, end, ref);

    const newComment: DocxComment = {
      id: commentId,
      author,
      ...(initials ? { initials } : {}),
      date: new Date(ctx.now()).toISOString(),
      body: [
        {
          kind: "paragraph",
          id: ctx.mintNodeId(),
          properties: {},
          children: [
            {
              kind: "run",
              id: ctx.mintNodeId(),
              properties: {},
              children: [textLeaf(ctx.mintNodeId, text)],
            } satisfies Run,
          ],
        } satisfies Paragraph,
      ],
    };

    const nextDoc = {
      ...replaceBlock(snapshot.root, idx, paragraphWithMarkers),
      comments: [...snapshot.root.comments, newComment],
    };
    const next = evolveSnapshot(snapshot, nextDoc, {
      body: true,
      comments: true,
      rels: true,
      contentTypes: true,
    });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: newComment.id,
        path: ["comments", snapshot.root.comments.length],
        summary: `+comment by ${author}: ${truncate(text, 40)}`,
      }),
    };
  },
};

function mintCommentId(existing: ReadonlyArray<DocxComment>): string {
  let max = -1;
  for (const c of existing) {
    const n = Number(c.id);
    if (Number.isFinite(n) && n > max) max = n;
  }
  return String(max + 1);
}

function wrapParagraphWithCommentMarkers(
  p: Paragraph,
  start: CommentRangeStart,
  end: CommentRangeEnd,
  ref: CommentReference
): Paragraph {
  // Single-paragraph case: prepend start marker, append end marker + reference
  // run. The reference must be inside a `<w:r>` per OOXML; the model emits a
  // synthetic run for it (handled in serializer).
  const children: InlineNode[] = [start, ...p.children, end, ref];
  return { ...p, children };
}

function truncate(s: string, n: number): string {
  if (s.length <= n) return s;
  return s.slice(0, n - 1) + "…";
}
