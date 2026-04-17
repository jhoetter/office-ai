import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxComment, DocxSnapshot, Paragraph, Run } from "../model/types.js";
import { buildDiff, evolveSnapshot, textLeaf } from "./helpers.js";
import type { ReplyCommentPayload } from "./payloads.js";

/**
 * Append a reply to an existing comment. The reply is itself a `w:comment`
 * with a fresh id; the parent relationship is encoded by `parentId`,
 * which the serializer materializes as a `w15:parentPaIdRef` entry in
 * `word/commentsExtended.xml`.
 *
 * Note: the reply does NOT add a new commentRangeStart/End/Reference
 * triple to the body. By OOXML convention, every reply in a thread shares
 * the parent's range markers — that's what makes Word render replies
 * indented under the same anchor in the comments pane.
 */
export const replyCommentHandler: CommandHandler<ReplyCommentPayload, DocxSnapshot> = {
  type: "docx:reply-comment",
  apply(snapshot, payload, ctx) {
    const { parentId, text, author, initials } = payload;
    const parentIdx = snapshot.root.comments.findIndex((c) => c.id === parentId);
    if (parentIdx < 0) {
      throw new CommandError("unknown-comment", `no parent comment with id "${parentId}"`);
    }
    if (!text || text.length === 0) {
      throw new CommandError("empty-reply", "reply-comment requires non-empty text");
    }

    const newId = mintCommentId(snapshot.root.comments);
    const replyParagraph: Paragraph = {
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
    };

    const reply: DocxComment = {
      id: newId,
      author,
      ...(initials ? { initials } : {}),
      date: new Date(ctx.now()).toISOString(),
      body: [replyParagraph],
      parentId,
    };

    const comments = [...snapshot.root.comments, reply];
    const next = evolveSnapshot(
      snapshot,
      { ...snapshot.root, comments },
      { comments: true, rels: true, contentTypes: true, commentsExtended: true }
    );
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: reply.id,
        path: ["comments", snapshot.root.comments.length],
        summary: `+reply by ${author} on comment ${parentId}: ${truncate(text, 40)}`,
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

function truncate(s: string, n: number): string {
  if (s.length <= n) return s;
  return s.slice(0, n - 1) + "…";
}
