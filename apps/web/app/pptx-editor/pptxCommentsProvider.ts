import type { CommentBody, CommentsProvider, CommentThread } from "@officeai/comments";
import { groupThreads } from "@officeai/comments";
import type { PptxAgent } from "@officeai/pptx/agent";
import type { PptxComment } from "@officeai/pptx";

/**
 * Adapter that exposes a `CommentsProvider` over a `PptxAgent`. The
 * shared `CommentsSidebar` (in `@officeai/ui`) drives PowerPoint
 * comments through this adapter, so the sidebar code itself doesn't
 * need to know anything about EMU pins or per-author idx counters.
 *
 * Each call snapshots the current agent state — there's no caching
 * here; React re-renders re-invoke `threads()` and we re-derive from
 * the latest `commentsByPart` map.
 */
export interface PptxCommentsProviderOptions {
  readonly agent: PptxAgent;
  readonly slideIndex: number;
  /** Editor-side hook that scrolls the canvas / flashes the pin. */
  readonly onScrollTo?: (commentId: string) => void;
  /** Default pin position for new comments; defaults to slide centre. */
  readonly defaultPin?: { xEmu: number; yEmu: number };
}

export function createPptxCommentsProvider(opts: PptxCommentsProviderOptions): CommentsProvider {
  const { agent, slideIndex } = opts;
  return {
    threads(): ReadonlyArray<CommentThread> {
      const snap = agent.getSnapshot();
      const slide = snap.root.slides[slideIndex];
      if (!slide || !slide.commentsPartPath) return [];
      const part = snap.root.commentsByPart.get(slide.commentsPartPath);
      if (!part) return [];
      const authorsById = new Map<number, string>();
      if (snap.root.commentAuthors) {
        for (const a of snap.root.commentAuthors.authors) authorsById.set(a.id, a.name);
      }
      const bodies: CommentBody[] = part.comments.map((c) => normalize(c, slideIndex, authorsById));
      return groupThreads(bodies);
    },
    async add(input) {
      const memo = await agent.applyCommand({
        type: "pptx:add-comment",
        payload: {
          slideIndex,
          author: input.author,
          text: input.text,
          ...(opts.defaultPin ? { xEmu: opts.defaultPin.xEmu, yEmu: opts.defaultPin.yEmu } : {}),
          ...(input.anchor.kind === "pptx-pin" ? { xEmu: input.anchor.xEmu, yEmu: input.anchor.yEmu } : {}),
          ...(input.anchor.kind === "pptx-pin" && input.anchor.shapeId
            ? { shapeId: input.anchor.shapeId }
            : {}),
        },
        source: "human",
      });
      return memoCommentId(memo, agent, slideIndex) ?? "";
    },
    async reply(input) {
      const memo = await agent.applyCommand({
        type: "pptx:reply-comment",
        payload: {
          slideIndex,
          parentId: input.parentId,
          author: input.author,
          text: input.text,
        },
        source: "human",
      });
      return memoCommentId(memo, agent, slideIndex) ?? "";
    },
    async resolve(commentId, resolved) {
      await agent.applyCommand({
        type: "pptx:resolve-comment",
        payload: { slideIndex, commentId, resolved },
        source: "human",
      });
    },
    async delete(commentId) {
      await agent.applyCommand({
        type: "pptx:delete-comment",
        payload: { slideIndex, commentId },
        source: "human",
      });
    },
    async edit(commentId, text) {
      await agent.applyCommand({
        type: "pptx:edit-comment",
        payload: { slideIndex, commentId, text },
        source: "human",
      });
    },
    onScrollTo: opts.onScrollTo,
  };
}

function normalize(
  c: PptxComment,
  slideIndex: number,
  authorsById: ReadonlyMap<number, string>
): CommentBody {
  return {
    id: c.id,
    author: authorsById.get(c.authorId) ?? `Author ${c.authorId}`,
    text: c.text,
    ...(c.createdAt ? { createdAt: c.createdAt } : {}),
    ...(c.resolved !== undefined ? { resolved: c.resolved } : {}),
    ...(c.parentId ? { parentId: c.parentId } : {}),
    anchor: {
      kind: "pptx-pin" as const,
      slideIndex,
      xEmu: c.xEmu,
      yEmu: c.yEmu,
    },
    nativeRef: c,
  };
}

function memoCommentId(_memo: unknown, agent: PptxAgent, slideIndex: number): string | null {
  // The command memo doesn't surface the new comment id directly; the
  // last entry on the slide's comments part is the freshest.
  const snap = agent.getSnapshot();
  const slide = snap.root.slides[slideIndex];
  if (!slide || !slide.commentsPartPath) return null;
  const part = snap.root.commentsByPart.get(slide.commentsPartPath);
  if (!part || part.comments.length === 0) return null;
  return part.comments[part.comments.length - 1]!.id;
}
