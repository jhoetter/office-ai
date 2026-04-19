import type { CommentBody, CommentsProvider, CommentThread } from "@officeai/comments";
import { groupThreads } from "@officeai/comments";
import type { XlsxAgent } from "@officeai/xlsx/agent";
import type { Comment as XlsxComment } from "@officeai/xlsx";

/**
 * Adapter that exposes a `CommentsProvider` over an `XlsxAgent`. The
 * shared `CommentsSidebar` (in `@officeai/ui`) drives Excel comments
 * through this adapter so the sidebar code itself doesn't have to
 * know about A1 refs or per-sheet comment parts.
 *
 * No caching: each render-time call to `threads()` re-derives from
 * the current agent snapshot, mirroring how the PPTX adapter works.
 */
export interface XlsxCommentsProviderOptions {
  readonly agent: XlsxAgent;
  /** Sheet name to scope comments to (sidebars are per-sheet). */
  readonly sheetName: string;
  /** Default A1 ref for new comments composed without a selection. */
  readonly defaultRef?: string;
  /** Editor-side hook that scrolls the grid to the cell. */
  readonly onScrollTo?: (commentId: string) => void;
}

export function createXlsxCommentsProvider(opts: XlsxCommentsProviderOptions): CommentsProvider {
  const { agent, sheetName } = opts;
  return {
    threads(): ReadonlyArray<CommentThread> {
      const sheet = agent.getSnapshot().root.sheets.find((s) => s.name === sheetName);
      if (!sheet) return [];
      const bodies: CommentBody[] = sheet.comments.map((c) => normalize(c, sheetName));
      return groupThreads(bodies);
    },
    async add(input) {
      const ref = input.anchor.kind === "xlsx-cell" ? input.anchor.ref : (opts.defaultRef ?? "A1");
      await agent.applyCommand({
        type: "xlsx:add-comment",
        payload: { sheet: sheetName, ref, author: input.author, text: input.text },
        source: "human",
      });
      return latestCommentId(agent, sheetName) ?? "";
    },
    async reply(input) {
      await agent.applyCommand({
        type: "xlsx:reply-comment",
        payload: {
          sheet: sheetName,
          parentId: input.parentId,
          author: input.author,
          text: input.text,
        },
        source: "human",
      });
      return latestCommentId(agent, sheetName) ?? "";
    },
    async resolve(commentId, resolved) {
      await agent.applyCommand({
        type: "xlsx:resolve-comment",
        payload: { sheet: sheetName, commentId, resolved },
        source: "human",
      });
    },
    async delete(commentId) {
      await agent.applyCommand({
        type: "xlsx:delete-comment",
        payload: { sheet: sheetName, commentId },
        source: "human",
      });
    },
    async edit(commentId, text) {
      await agent.applyCommand({
        type: "xlsx:edit-comment",
        payload: { sheet: sheetName, commentId, text },
        source: "human",
      });
    },
    onScrollTo: opts.onScrollTo,
  };
}

function normalize(c: XlsxComment, sheetName: string): CommentBody {
  return {
    id: c.id,
    author: c.author,
    text: c.text,
    ...(c.createdAt ? { createdAt: c.createdAt } : {}),
    ...(c.resolved !== undefined ? { resolved: c.resolved } : {}),
    ...(c.parentId ? { parentId: c.parentId } : {}),
    anchor: {
      kind: "xlsx-cell" as const,
      sheet: sheetName,
      ref: c.ref,
    },
    nativeRef: c,
  };
}

function latestCommentId(agent: XlsxAgent, sheetName: string): string | null {
  const sheet = agent.getSnapshot().root.sheets.find((s) => s.name === sheetName);
  if (!sheet || sheet.comments.length === 0) return null;
  return sheet.comments[sheet.comments.length - 1].id;
}
