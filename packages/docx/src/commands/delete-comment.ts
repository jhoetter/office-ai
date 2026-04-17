import { CommandError, type CommandHandler } from "@officeai/core";
import type {
  BlockNode,
  DocxComment,
  DocxDocument,
  DocxSnapshot,
  InlineNode,
  Paragraph,
} from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { DeleteCommentPayload } from "./payloads.js";

/**
 * Drop a comment from the document along with its inline range markers.
 *
 * Threading semantics: deleting a parent comment also deletes every reply
 * whose `parentId` chains back to it (transitive). We picked "delete the
 * thread" over "refuse with thread-non-empty" because (a) it matches Word's
 * behavior when a user deletes the head of a thread and (b) keeps the
 * command idempotent — re-running on an already-empty subtree does nothing.
 */
export const deleteCommentHandler: CommandHandler<DeleteCommentPayload, DocxSnapshot> = {
  type: "docx:delete-comment",
  apply(snapshot, payload) {
    const { commentId } = payload;
    const exists = snapshot.root.comments.some((c) => c.id === commentId);
    if (!exists) {
      throw new CommandError("unknown-comment", `no comment with id "${commentId}"`);
    }

    const toRemove = collectThread(snapshot.root.comments, commentId);
    const removeSet = new Set(toRemove);
    const remainingComments = snapshot.root.comments.filter((c) => !removeSet.has(c.id));
    const newBody = stripCommentMarkers(snapshot.root.body, removeSet);

    const willHaveCommentsExtended = remainingComments.some(
      (c) => c.resolved === true || c.parentId !== undefined
    );

    const nextDoc: DocxDocument = {
      ...snapshot.root,
      body: newBody,
      comments: remainingComments,
    };
    const next = evolveSnapshot(snapshot, nextDoc, {
      body: true,
      comments: true,
      rels: remainingComments.length === 0,
      contentTypes: remainingComments.length === 0,
      commentsExtended: willHaveCommentsExtended || snapshot.dirty.commentsExtended,
    });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-deleted",
        nodeId: commentId,
        path: ["comments"],
        summary:
          toRemove.length === 1
            ? `−comment ${commentId}`
            : `−comment thread (${toRemove.length}: ${toRemove.join(", ")})`,
      }),
    };
  },
};

function collectThread(comments: ReadonlyArray<DocxComment>, root: string): string[] {
  const out: string[] = [root];
  const queue: string[] = [root];
  while (queue.length > 0) {
    const parent = queue.shift() as string;
    for (const c of comments) {
      if (c.parentId === parent && !out.includes(c.id)) {
        out.push(c.id);
        queue.push(c.id);
      }
    }
  }
  return out;
}

function stripCommentMarkers(body: ReadonlyArray<BlockNode>, removeIds: Set<string>): BlockNode[] {
  return body.map((b) => (b.kind === "paragraph" ? stripFromParagraph(b, removeIds) : b));
}

function stripFromParagraph(p: Paragraph, removeIds: Set<string>): Paragraph {
  const children: InlineNode[] = [];
  for (const child of p.children) {
    if (
      (child.kind === "comment-range-start" ||
        child.kind === "comment-range-end" ||
        child.kind === "comment-reference") &&
      removeIds.has(child.commentId)
    ) {
      continue;
    }
    children.push(child);
  }
  if (children.length === p.children.length) return p;
  return { ...p, children };
}
