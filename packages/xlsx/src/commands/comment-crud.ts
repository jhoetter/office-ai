/**
 * P1 Comment CRUD: reply / resolve / delete / edit.
 *
 * Each handler mutates the typed `Sheet.comments` list and dirties the
 * underlying `xl/comments{N}.xml` part. Threaded comments persist as
 * additional `<comment>` entries with a `parentId` field — Excel
 * round-trips them via its own `xl/threadedComments/*` parts when
 * loaded by a modern client; legacy clients still see the parent + the
 * replies as flat notes (acceptable for the headless surface).
 *
 * Spec: `spec/xlsx/agent-commands.md` §13.
 */

import { CommandError, type CommandHandler, type DiffChange } from "@officeai/core";
import type { Comment, Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type {
  DeleteCommentPayload,
  EditCommentPayload,
  ReplyCommentPayload,
  ResolveCommentPayload,
} from "./payloads.js";
import { resolveSheet } from "./validation.js";

export const replyCommentHandler: CommandHandler<ReplyCommentPayload, XlsxSnapshot> = {
  type: "xlsx:reply-comment",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    requireString(payload.author, "Author is required");
    requireString(payload.text, "Comment text must be non-empty");
    const parent = sheet.comments.find((c) => c.id === payload.parentId);
    if (!parent) {
      throw new CommandError(
        "unknown-target",
        `No comment with id ${JSON.stringify(payload.parentId)} on sheet ${sheet.name}`
      );
    }
    if (parent.parentId) {
      throw new CommandError(
        "invalid-payload",
        "Cannot reply to a reply — replies must target a top-level comment"
      );
    }
    const commentId = mintCommentId(sheet);
    const commentAuthors = sheet.commentAuthors.includes(payload.author)
      ? sheet.commentAuthors
      : [...sheet.commentAuthors, payload.author];
    const reply: Comment = {
      id: commentId,
      ref: parent.ref,
      author: payload.author,
      text: payload.text,
      parentId: parent.id,
      createdAt: new Date().toISOString(),
    };
    const nextSheet: Sheet = {
      ...sheet,
      comments: [...sheet.comments, reply],
      commentAuthors,
    };
    return commit(snapshot, sheet, nextSheet, [
      {
        kind: "node-inserted",
        nodeId: nextSheet.id,
        path: ["sheets", nextSheet.index, "comments", commentId],
        summary: `Replied to ${payload.parentId} on ${sheet.name}`,
        meta: { commentId, parentId: parent.id, author: payload.author },
      },
    ]);
  },
};

export const resolveCommentHandler: CommandHandler<ResolveCommentPayload, XlsxSnapshot> = {
  type: "xlsx:resolve-comment",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const target = sheet.comments.find((c) => c.id === payload.commentId);
    if (!target) {
      throw new CommandError(
        "unknown-target",
        `No comment with id ${JSON.stringify(payload.commentId)} on sheet ${sheet.name}`
      );
    }
    if (target.parentId) {
      throw new CommandError(
        "invalid-payload",
        "Resolve only applies to top-level comments — resolve the parent thread instead"
      );
    }
    const nextSheet: Sheet = {
      ...sheet,
      comments: sheet.comments.map((c) =>
        c.id === payload.commentId ? { ...c, resolved: payload.resolved } : c
      ),
    };
    return commit(snapshot, sheet, nextSheet, [
      {
        kind: "node-updated",
        nodeId: nextSheet.id,
        path: ["sheets", nextSheet.index, "comments", payload.commentId],
        field: "resolved",
        summary: payload.resolved ? "resolved" : "reopened",
      },
    ]);
  },
};

export const deleteCommentHandler: CommandHandler<DeleteCommentPayload, XlsxSnapshot> = {
  type: "xlsx:delete-comment",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const target = sheet.comments.find((c) => c.id === payload.commentId);
    if (!target) {
      throw new CommandError(
        "unknown-target",
        `No comment with id ${JSON.stringify(payload.commentId)} on sheet ${sheet.name}`
      );
    }
    const remaining = sheet.comments.filter((c) => {
      if (c.id === payload.commentId) return false;
      // Cascade: dropping a top-level comment removes its replies too.
      if (!target.parentId && c.parentId === payload.commentId) return false;
      return true;
    });
    const nextSheet: Sheet = { ...sheet, comments: remaining };
    return commit(snapshot, sheet, nextSheet, [
      {
        kind: "node-deleted",
        nodeId: nextSheet.id,
        path: ["sheets", nextSheet.index, "comments", payload.commentId],
        summary: `Deleted ${payload.commentId} on ${sheet.name}`,
      },
    ]);
  },
};

export const editCommentHandler: CommandHandler<EditCommentPayload, XlsxSnapshot> = {
  type: "xlsx:edit-comment",
  apply(snapshot, payload) {
    if (typeof payload.text !== "string") {
      throw new CommandError("invalid-payload", "text must be a string");
    }
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const target = sheet.comments.find((c) => c.id === payload.commentId);
    if (!target) {
      throw new CommandError(
        "unknown-target",
        `No comment with id ${JSON.stringify(payload.commentId)} on sheet ${sheet.name}`
      );
    }
    const nextSheet: Sheet = {
      ...sheet,
      comments: sheet.comments.map((c) =>
        c.id === payload.commentId ? { ...c, text: payload.text } : c
      ),
    };
    return commit(snapshot, sheet, nextSheet, [
      {
        kind: "node-updated",
        nodeId: nextSheet.id,
        path: ["sheets", nextSheet.index, "comments", payload.commentId],
        field: "text",
        summary: `text:${payload.text.length}ch`,
      },
    ]);
  },
};

// ─── Internals ────────────────────────────────────────────────────────────

function commit(
  snapshot: XlsxSnapshot,
  prevSheet: Sheet,
  nextSheet: Sheet,
  changes: DiffChange[]
) {
  const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
  const dirtyPatch: Parameters<typeof evolveSnapshot>[2] = {};
  if (prevSheet.commentsPartPath) {
    dirtyPatch.comments = [prevSheet.commentsPartPath];
  }
  const next = evolveSnapshot(snapshot, nextWorkbook, dirtyPatch);
  return {
    next,
    diff: buildDiff(snapshot.revision, next.revision, changes),
  };
}

function mintCommentId(sheet: Sheet): string {
  // Use the largest existing numeric suffix + 1 so cascading deletes
  // don't re-issue an id that's already been seen by collaborators.
  let max = 0;
  for (const c of sheet.comments) {
    const m = /^comment-(\d+)$/.exec(c.id);
    if (m) {
      const n = Number(m[1]);
      if (n > max) max = n;
    }
  }
  return `comment-${max + 1}`;
}

function requireString(value: string, message: string): void {
  if (typeof value !== "string" || value.length === 0) {
    throw new CommandError("invalid-payload", message);
  }
}
