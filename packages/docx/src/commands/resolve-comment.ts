import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxComment, DocxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { ResolveCommentPayload } from "./payloads.js";

/**
 * Toggle a comment's resolved state. Drives `word/commentsExtended.xml`
 * (`w15:commentEx[@w15:done]`). Idempotent: setting the current state
 * again is treated as a no-op but still bumps the revision so observers
 * (and the bus's history) record an explicit "user attempted" event.
 */
export const resolveCommentHandler: CommandHandler<ResolveCommentPayload, DocxSnapshot> = {
  type: "docx:resolve-comment",
  apply(snapshot, payload) {
    const { commentId } = payload;
    const targetResolved = payload.resolved ?? true;
    const idx = snapshot.root.comments.findIndex((c) => c.id === commentId);
    if (idx < 0) {
      throw new CommandError("unknown-comment", `no comment with id "${commentId}"`);
    }
    const current = snapshot.root.comments[idx];
    const isCurrentlyResolved = current.resolved === true;
    if (isCurrentlyResolved === targetResolved) {
      const next = evolveSnapshot(snapshot, snapshot.root, {});
      return {
        next,
        diff: buildDiff(snapshot.revision, next.revision, {
          kind: "node-updated",
          nodeId: current.id,
          path: ["comments", idx],
          field: "resolved",
          summary: `comment ${commentId} already ${targetResolved ? "resolved" : "open"} (no-op)`,
        }),
      };
    }
    const updated: DocxComment = targetResolved ? { ...current, resolved: true } : stripResolved(current);
    const comments = snapshot.root.comments.slice();
    comments[idx] = updated;
    const next = evolveSnapshot(snapshot, { ...snapshot.root, comments }, { commentsExtended: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: current.id,
        path: ["comments", idx],
        field: "resolved",
        summary: targetResolved ? `comment ${commentId} resolved` : `comment ${commentId} reopened`,
      }),
    };
  },
};

function stripResolved(c: DocxComment): DocxComment {
  const { resolved: _ignored, ...rest } = c;
  void _ignored;
  return rest;
}
