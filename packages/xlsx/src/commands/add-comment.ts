import { CommandError, ooxml, type CommandHandler, type DiffChange } from "@officeai/core";
import type { Comment, Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { AddCommentPayload } from "./payloads.js";
import { parseCellRef, resolveSheet } from "./validation.js";

/**
 * `xlsx:add-comment` — attach a classic note to a single cell.
 *
 * Spec: `spec/xlsx/agent-commands.md` §13.
 *
 * Pipeline:
 *   1. Resolve the sheet + ref. Range refs are rejected (`invalid-ref`).
 *   2. Validate non-empty text and author.
 *   3. Reject `comment-exists` if the cell already carries a comment.
 *   4. Append the author to `commentAuthors` if missing (de-duped by
 *      string equality, preserving insertion order to keep
 *      `authorId` indices stable across the round-trip).
 *   5. Mint a stable `commentId` based on the sheet's running counter.
 *   6. If the sheet has no comments part yet, mint a fresh
 *      `xl/comments{N}.xml` path that doesn't collide with any other
 *      sheet's comments part or any existing container part. Dirty
 *      `sheetRels` (so the per-sheet rels file gains the comments
 *      relationship) and `contentTypes` (so the override is added).
 *   7. Always dirty the comments part itself.
 *
 * Out of scope for P0 (deferred to P1, documented in
 * `docs/build-log/xlsx.md` Phase 7j):
 *   - Threaded comments (`xl/threadedComments/*`) — kept opaque.
 *   - VML drawings (the legacy `<v:shape>` markup that anchors
 *     comments visually). Without VML, the comment round-trips
 *     correctly in the data layer but won't render with a pinned
 *     position in Excel. Acceptable for the headless P0 surface.
 */
export const addCommentHandler: CommandHandler<AddCommentPayload, XlsxSnapshot> = {
  type: "xlsx:add-comment",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);

    if (payload.ref.includes(":")) {
      throw new CommandError(
        "invalid-ref",
        `Comment ref ${JSON.stringify(payload.ref)} must be a single A1 cell, not a range`
      );
    }
    const addr = parseCellRef(payload.ref);
    void addr;

    if (payload.text === "") {
      throw new CommandError("empty-text", "Comment text must be non-empty");
    }
    if (payload.author === "") {
      throw new CommandError("empty-author", "Author is required");
    }

    if (sheet.comments.find((c) => c.ref === payload.ref)) {
      throw new CommandError(
        "comment-exists",
        `Cell ${payload.sheet}!${payload.ref} already has a comment; call xlsx:delete-comment first then re-add`
      );
    }

    const commentAuthors = sheet.commentAuthors.includes(payload.author)
      ? sheet.commentAuthors
      : [...sheet.commentAuthors, payload.author];

    const commentId = `comment-${sheet.comments.length + 1}`;
    const newComment: Comment = {
      id: commentId,
      ref: payload.ref,
      author: payload.author,
      text: payload.text,
    };
    const comments = [...sheet.comments, newComment];

    const isFirstComment = sheet.commentsPartPath === undefined;
    const commentsPartPath = isFirstComment ? mintCommentsPartPath(snapshot) : sheet.commentsPartPath!;

    const nextSheet: Sheet = {
      ...sheet,
      comments,
      commentAuthors,
      ...(isFirstComment ? { commentsPartPath } : {}),
    };
    const nextWorkbook = replaceSheet(snapshot.root, nextSheet);

    const dirtyPatch: Parameters<typeof evolveSnapshot>[2] = {
      comments: [commentsPartPath],
    };
    if (isFirstComment) {
      dirtyPatch.sheetRels = [ooxml.RelationshipGraph.relsPathFor(sheet.partPath)];
      dirtyPatch.contentTypes = true;
    }
    const next = evolveSnapshot(snapshot, nextWorkbook, dirtyPatch);

    const changes: DiffChange[] = [
      {
        kind: "node-inserted",
        nodeId: nextSheet.id,
        path: ["sheets", nextSheet.index, "comments", commentId],
        summary: `Added comment by ${payload.author} on ${sheet.name}!${payload.ref}`,
        meta: {
          sheet: sheet.name,
          ref: payload.ref,
          commentId,
          author: payload.author,
          text: payload.text,
        },
      },
    ];
    if (isFirstComment) {
      changes.push({
        kind: "node-inserted",
        nodeId: nextSheet.id,
        path: ["parts", commentsPartPath],
        summary: `Created ${commentsPartPath}`,
        meta: { partPath: commentsPartPath },
      });
    }

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, changes),
    };
  },
};

/**
 * Pick a fresh `xl/comments{N}.xml` path. `N` is the smallest positive
 * integer not already claimed by any sheet's `commentsPartPath` and
 * not already present as a container key — covers both the "first
 * comments part in the workbook" case and the "sheet picked up an
 * existing-but-unbound xl/commentsK.xml at parse time" case (which
 * shouldn't happen in practice but is cheap to defend against).
 */
function mintCommentsPartPath(snapshot: XlsxSnapshot): string {
  const taken = new Set<string>();
  for (const s of snapshot.root.sheets) {
    if (s.commentsPartPath) taken.add(s.commentsPartPath);
  }
  for (const path of snapshot.container.parts.keys()) taken.add(path);
  let i = 1;
  while (taken.has(`xl/comments${i}.xml`)) i++;
  return `xl/comments${i}.xml`;
}
