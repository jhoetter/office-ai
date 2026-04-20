"use client";

import { useMemo } from "react";
import type { CommentBody, CommentsProvider, CommentThread } from "@officeai/comments";
import { groupThreads } from "@officeai/comments";
import type { PdfAgent } from "@officeai/pdf/agent";
import type { PdfComment } from "@officeai/pdf";

/**
 * Bridge between the headless `PdfAgent.snapshot.root.comments` array
 * and the shared `CommentsSidebar` React component (in
 * `@officeai/ui`).
 *
 * Each PDF comment carries:
 *   - `pageNumber` (1-indexed)
 *   - `normalizedRect` ([0..1, 0..1, 0..1, 0..1] in MediaBox space)
 *
 * which we project into the cross-format `pdf-region` anchor variant
 * that `@officeai/comments` already understands. The `PdfCanvas` reads
 * that anchor back when the user clicks a comment in the sidebar to
 * scroll the page into view and flash the rect.
 *
 * The provider is a pure function of `(agent, currentPage)`; we do
 * not memoize internally — the `CommentsSidebar` re-renders on every
 * snapshot change, which is exactly the cadence we want for a fresh
 * `threads()` re-derivation. Wrap it in `useMemo` at the call site if
 * you need referential stability.
 */
export interface PdfCommentsProviderOptions {
  readonly agent: PdfAgent;
  /** 1-indexed page number used as the default anchor for newly added
   * comments. The PDF canvas passes the page currently in the centre
   * of the viewport. */
  readonly currentPage: number;
  /** Optional editor-side hook to scroll the canvas to a clicked
   * comment's anchor. The shared sidebar calls it via
   * `provider.onScrollTo`. */
  readonly onScrollTo?: (commentId: string) => void;
  /** Default normalized rect for new comments. Defaults to a small
   * pin in the top-left corner. */
  readonly defaultRect?: readonly [number, number, number, number];
  /** Author display name written into new comment bodies. Defaults to
   * "You" — matches the PPTX behaviour. */
  readonly author?: string;
}

const DEFAULT_RECT: readonly [number, number, number, number] = [0.05, 0.92, 0.13, 0.96];

export function createPdfCommentsProvider(opts: PdfCommentsProviderOptions): CommentsProvider {
  const { agent, currentPage } = opts;
  const author = opts.author ?? "You";
  return {
    threads(): ReadonlyArray<CommentThread> {
      const snap = agent.getSnapshot();
      const bodies: CommentBody[] = snap.root.comments.map((c) => normalizeComment(c));
      return groupThreads(bodies);
    },
    async add(input) {
      const anchor = input.anchor;
      const rect: readonly [number, number, number, number] =
        anchor.kind === "pdf-region"
          ? anchor.normalizedRect
          : (opts.defaultRect ?? DEFAULT_RECT);
      const pageNumber = anchor.kind === "pdf-region" ? anchor.pageNumber : currentPage;
      const before = agent.getSnapshot().root.comments;
      await agent.applyCommand({
        type: "pdf:add-comment",
        payload: {
          author: input.author,
          text: input.text,
          pageNumber,
          normalizedRect: rect,
        },
        source: "human",
      });
      return latestNewCommentId(agent, before) ?? "";
    },
    async reply(input) {
      const before = agent.getSnapshot().root.comments;
      await agent.applyCommand({
        type: "pdf:reply-comment",
        payload: {
          parentId: input.parentId,
          author: input.author,
          text: input.text,
        },
        source: "human",
      });
      return latestNewCommentId(agent, before) ?? "";
    },
    async resolve(commentId, resolved) {
      await agent.applyCommand({
        type: "pdf:resolve-comment",
        payload: { commentId, resolved },
        source: "human",
      });
    },
    async delete(commentId) {
      await agent.applyCommand({
        type: "pdf:delete-comment",
        payload: { commentId },
        source: "human",
      });
    },
    async edit(commentId, text) {
      await agent.applyCommand({
        type: "pdf:edit-comment",
        payload: { commentId, text },
        source: "human",
      });
    },
    onScrollTo: opts.onScrollTo,
  };
}

/**
 * React hook wrapper. Memoizes the provider on agent identity +
 * currentPage so React's `===` checks in the shared sidebar don't
 * tear down per render. Also re-uses the same `onScrollTo` callback
 * so the sidebar's effect dependencies stay stable.
 */
export function usePdfCommentsProvider(opts: PdfCommentsProviderOptions): CommentsProvider {
  const { agent, currentPage, onScrollTo, author } = opts;
  return useMemo<CommentsProvider>(
    () =>
      createPdfCommentsProvider({
        agent,
        currentPage,
        ...(onScrollTo ? { onScrollTo } : {}),
        ...(author ? { author } : {}),
      }),
    [agent, currentPage, onScrollTo, author]
  );
}

function normalizeComment(c: PdfComment): CommentBody {
  return {
    id: c.id,
    author: c.author,
    text: c.text,
    ...(c.createdAt ? { createdAt: c.createdAt } : {}),
    ...(c.resolved !== undefined ? { resolved: c.resolved } : {}),
    ...(c.parentId ? { parentId: c.parentId } : {}),
    anchor: {
      kind: "pdf-region" as const,
      pageNumber: c.pageNumber,
      normalizedRect: c.normalizedRect,
    },
    nativeRef: c,
  };
}

function latestNewCommentId(
  agent: PdfAgent,
  before: ReadonlyArray<PdfComment>
): string | null {
  const after = agent.getSnapshot().root.comments;
  if (after.length === before.length) return null;
  const beforeIds = new Set(before.map((c) => c.id));
  for (let i = after.length - 1; i >= 0; i -= 1) {
    const c = after[i];
    if (c && !beforeIds.has(c.id)) return c.id;
  }
  return null;
}
