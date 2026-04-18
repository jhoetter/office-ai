import type { CommentBody, CommentThread } from "./types";

/**
 * Group an unordered list of comments into threads. Top-level comments
 * (no `parentId`) become thread parents in their original order;
 * comments with a `parentId` are appended to their parent's reply list.
 *
 * Orphan replies — `parentId` pointing at a comment that doesn't exist
 * (or is itself a reply with no resolved root) — are surfaced as their
 * own top-level threads instead of being silently dropped, so the UI
 * can still show them and the user can clean them up.
 *
 * Stable: input order determines parent order; reply order matches
 * input order within each parent.
 */
export function groupThreads(comments: ReadonlyArray<CommentBody>): ReadonlyArray<CommentThread> {
  const byId = new Map<string, CommentBody>();
  for (const c of comments) byId.set(c.id, c);

  const parents: { parent: CommentBody; replies: CommentBody[] }[] = [];
  const parentIndex = new Map<string, { parent: CommentBody; replies: CommentBody[] }>();
  for (const c of comments) {
    if (!c.parentId) {
      const t = { parent: c, replies: [] };
      parents.push(t);
      parentIndex.set(c.id, t);
    }
  }

  const orphans: { parent: CommentBody; replies: CommentBody[] }[] = [];
  for (const c of comments) {
    if (!c.parentId) continue;
    const parent = byId.get(c.parentId);
    if (parent && !parent.parentId && parentIndex.has(parent.id)) {
      parentIndex.get(parent.id)!.replies.push(c);
    } else {
      orphans.push({ parent: c, replies: [] });
    }
  }

  return [...parents, ...orphans];
}

/**
 * Count unresolved threads — useful for an inbox-style "N open
 * comments" badge in the toolbar.
 */
export function countOpenThreads(threads: ReadonlyArray<CommentThread>): number {
  let n = 0;
  for (const t of threads) if (!t.parent.resolved) n++;
  return n;
}
