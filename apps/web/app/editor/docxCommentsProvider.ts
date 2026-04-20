import type { CommentBody, CommentsProvider, CommentThread } from "@officeai/comments";
import { groupThreads } from "@officeai/comments";
import type { DocxAgent, DocxComment, DocxSelection } from "@officeai/docx";
import { commentParagraphIndex, commentPlainText } from "@/lib/format-helpers";

/**
 * Adapter that exposes a `CommentsProvider` over a `DocxAgent`. The
 * shared `CommentsSidebar` (in `@officeai/ui`) drives Word comments
 * through this adapter, mirroring the PPTX / XLSX setups so all three
 * editors share one sidebar surface.
 *
 * Two DOCX-specific tricks live here:
 *   - `add` requires a `DocxSelection` range (you can't comment without
 *     selecting text); we lift it off the `docx-range` anchor's opaque
 *     `range` field, which the editor populates from the live PM
 *     selection at composer-open time.
 *   - The DOCX comment body is rich (Block[]) but the shared UI only
 *     surfaces plain text — `commentPlainText` flattens it and the
 *     adapter stashes the original `DocxComment` on `nativeRef` so any
 *     future format-aware view can recover it.
 */
export interface DocxCommentsProviderOptions {
  readonly agent: DocxAgent;
  /** Called when the sidebar selects a thread; scrolls + flashes the comment. */
  readonly onScrollTo?: (commentId: string) => void;
  /** Default author/initials for new comments; defaults to "You" / "Y". */
  readonly defaultAuthor?: { name: string; initials?: string };
  /**
   * Optional toast hook so the editor surface keeps the lightweight
   * confirmation feedback ("Reply added.", "Comment resolved.", …) it
   * had before the migration. Adapter callers (CLI, tests) leave it
   * undefined.
   */
  readonly onToast?: (kind: "info" | "warn" | "error", text: string) => void;
}

export function createDocxCommentsProvider(opts: DocxCommentsProviderOptions): CommentsProvider {
  const { agent } = opts;
  const author = opts.defaultAuthor ?? { name: "You", initials: "Y" };
  const toast = opts.onToast;
  // OOXML doesn't carry our realtime peer id / color — only an
  // author display name. We stash the realtime identity for newly
  // authored comments here so `normalize()` can hydrate
  // `authorId` / `authorColor` on the next `threads()` snapshot.
  // Cleared implicitly when the document is reloaded (the provider
  // is recreated alongside the new agent).
  const localIdentity = new Map<string, { readonly authorId?: string; readonly authorColor?: string }>();
  const stamp = (id: string, input: { authorId?: string; authorColor?: string }): void => {
    if (!input.authorId && !input.authorColor) return;
    const next: { authorId?: string; authorColor?: string } = {};
    if (input.authorId) next.authorId = input.authorId;
    if (input.authorColor) next.authorColor = input.authorColor;
    localIdentity.set(id, next);
  };
  const toast_ = toast;
  const guarded = async <T>(label: string, op: () => Promise<T>): Promise<T | undefined> => {
    try {
      const result = await op();
      if (toast_ && label) toast_("info", label);
      return result;
    } catch (err) {
      if (toast_) toast_("error", err instanceof Error ? err.message : String(err));
      return undefined;
    }
  };
  return {
    threads(): ReadonlyArray<CommentThread> {
      const snap = agent.getSnapshot();
      const bodies: CommentBody[] = snap.root.comments.map((c) => {
        const base = normalize(c, snap);
        const stamped = localIdentity.get(c.id);
        if (!stamped) return base;
        return { ...base, ...stamped };
      });
      return groupThreads(bodies);
    },
    async add(input) {
      const range =
        input.anchor.kind === "docx-range" ? (input.anchor.range as DocxSelection | undefined) : undefined;
      if (!range) {
        throw new Error(
          "DocxCommentsProvider.add(): the docx-range anchor must carry an opaque DocxSelection range"
        );
      }
      await guarded("Comment added.", () =>
        agent.applyCommand({
          type: "docx:add-comment",
          payload: {
            range,
            text: input.text,
            author: input.author || author.name,
            ...(author.initials ? { initials: author.initials } : {}),
          },
          source: "human",
        })
      );
      const newId = lastCommentId(agent) ?? "";
      if (newId) stamp(newId, input);
      return newId;
    },
    async reply(input) {
      await guarded("Reply added.", () =>
        agent.applyCommand({
          type: "docx:reply-comment",
          payload: {
            parentId: input.parentId,
            text: input.text,
            author: input.author || author.name,
            ...(author.initials ? { initials: author.initials } : {}),
          },
          source: "human",
        })
      );
      const newId = lastCommentId(agent) ?? "";
      if (newId) stamp(newId, input);
      return newId;
    },
    async resolve(commentId, resolved) {
      await guarded(resolved ? "Comment resolved." : "Comment reopened.", () =>
        agent.applyCommand({
          type: "docx:resolve-comment",
          payload: { commentId, resolved },
          source: "human",
        })
      );
    },
    async delete(commentId) {
      await guarded("Comment deleted.", () =>
        agent.applyCommand({
          type: "docx:delete-comment",
          payload: { commentId },
          source: "human",
        })
      );
    },
    onScrollTo: opts.onScrollTo,
  };
}

function normalize(c: DocxComment, snap: ReturnType<DocxAgent["getSnapshot"]>): CommentBody {
  const paragraphIndex = commentParagraphIndex(snap, c.id) ?? 0;
  return {
    id: c.id,
    author: c.author,
    text: commentPlainText(c),
    ...(c.date ? { createdAt: c.date } : {}),
    ...(c.resolved !== undefined ? { resolved: c.resolved } : {}),
    ...(c.parentId ? { parentId: c.parentId } : {}),
    anchor: {
      kind: "docx-range" as const,
      paragraphIndex,
    },
    nativeRef: c,
  };
}

function lastCommentId(agent: DocxAgent): string | null {
  const list = agent.getSnapshot().root.comments;
  if (list.length === 0) return null;
  return list[list.length - 1].id;
}
