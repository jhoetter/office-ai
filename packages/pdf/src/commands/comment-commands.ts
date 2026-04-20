import type { CommandHandler } from "@officeai/core";
import type { PdfComment, PdfDocument, PdfSnapshot } from "../model/types.js";
import { buildDiff, evolvePdf, makeError } from "./helpers.js";
import type {
  AddCommentPayload,
  DeleteCommentPayload,
  EditCommentPayload,
  ReplyCommentPayload,
  ResolveCommentPayload,
} from "./payloads.js";

const cloneComments = (snapshot: PdfSnapshot): PdfComment[] => [...snapshot.root.comments];

const findIndex = (comments: ReadonlyArray<PdfComment>, id: string): number => {
  const idx = comments.findIndex((c) => c.id === id);
  if (idx < 0) throw makeError("unknown-target", `comment ${id} not found`);
  return idx;
};

export const addCommentHandler: CommandHandler<AddCommentPayload, PdfSnapshot> = {
  type: "pdf:add-comment",
  apply(snapshot, payload, ctx) {
    const id = payload.id ?? ctx.mintNodeId();
    const comment: PdfComment = {
      id,
      author: payload.author,
      text: payload.text,
      pageNumber: payload.pageNumber,
      normalizedRect: payload.normalizedRect,
      createdAt: new Date(ctx.now()).toISOString(),
    };
    const comments = [...snapshot.root.comments, comment];
    const root: PdfDocument = { ...snapshot.root, comments };
    const next = evolvePdf(snapshot, root);
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: id,
        path: ["comments", comments.length - 1],
        summary: "comment",
      }),
    };
  },
};

export const replyCommentHandler: CommandHandler<ReplyCommentPayload, PdfSnapshot> = {
  type: "pdf:reply-comment",
  apply(snapshot, payload, ctx) {
    const parent = snapshot.root.comments.find((c) => c.id === payload.parentId);
    if (!parent) throw makeError("unknown-target", `comment ${payload.parentId} not found`);
    const id = payload.id ?? ctx.mintNodeId();
    const reply: PdfComment = {
      id,
      author: payload.author,
      text: payload.text,
      parentId: payload.parentId,
      pageNumber: parent.pageNumber,
      normalizedRect: parent.normalizedRect,
      createdAt: new Date(ctx.now()).toISOString(),
    };
    const comments = [...snapshot.root.comments, reply];
    const root: PdfDocument = { ...snapshot.root, comments };
    const next = evolvePdf(snapshot, root);
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: id,
        path: ["comments", comments.length - 1],
        summary: "comment-reply",
      }),
    };
  },
};

export const editCommentHandler: CommandHandler<EditCommentPayload, PdfSnapshot> = {
  type: "pdf:edit-comment",
  apply(snapshot, payload) {
    const comments = cloneComments(snapshot);
    const idx = findIndex(comments, payload.commentId);
    comments[idx] = { ...comments[idx], text: payload.text };
    const root: PdfDocument = { ...snapshot.root, comments };
    const next = evolvePdf(snapshot, root);
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: payload.commentId,
        path: ["comments", idx, "text"],
        field: "text",
        summary: "comment-text",
      }),
    };
  },
};

export const resolveCommentHandler: CommandHandler<ResolveCommentPayload, PdfSnapshot> = {
  type: "pdf:resolve-comment",
  apply(snapshot, payload) {
    const comments = cloneComments(snapshot);
    const idx = findIndex(comments, payload.commentId);
    comments[idx] = { ...comments[idx], resolved: payload.resolved };
    const root: PdfDocument = { ...snapshot.root, comments };
    const next = evolvePdf(snapshot, root);
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: payload.commentId,
        path: ["comments", idx, "resolved"],
        field: "resolved",
        summary: "comment-resolved",
      }),
    };
  },
};

export const deleteCommentHandler: CommandHandler<DeleteCommentPayload, PdfSnapshot> = {
  type: "pdf:delete-comment",
  apply(snapshot, payload) {
    const comments = cloneComments(snapshot);
    const idx = findIndex(comments, payload.commentId);
    const [removed] = comments.splice(idx, 1);
    const repliesRemoved = comments.length;
    for (let i = comments.length - 1; i >= 0; i--) {
      if (comments[i].parentId === payload.commentId) comments.splice(i, 1);
    }
    const root: PdfDocument = { ...snapshot.root, comments };
    const next = evolvePdf(snapshot, root);
    void repliesRemoved;
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-deleted",
        nodeId: removed.id,
        path: ["comments", idx],
        summary: "comment",
      }),
    };
  },
};
