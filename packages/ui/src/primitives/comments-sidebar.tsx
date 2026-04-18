"use client";

import { useState, type FormEvent, type ReactNode } from "react";
import { Check, CornerDownRight, MessageCircle, Trash2 } from "lucide-react";
import { groupThreads, type CommentBody, type CommentsProvider, type CommentThread } from "@officeai/comments";
import { cn } from "../lib/cn";
import { Button } from "./button";

export interface CommentsSidebarProps {
  /**
   * Either a `CommentsProvider` (preferred — adapter style for new
   * editors) or a pre-grouped list of threads (escape hatch for
   * editors that haven't been migrated yet).
   */
  provider?: CommentsProvider;
  threads?: ReadonlyArray<CommentThread>;
  /** Author name to attach to new replies; defaults to `"You"`. */
  author?: string;
  /** Editor-specific scroll-to-anchor side-effect. */
  onScrollTo?: (commentId: string) => void;
  /** Empty-state copy. Defaults to a generic "no comments yet" line. */
  emptyHint?: ReactNode;
}

/**
 * Format-agnostic comments sidebar. Lists every top-level comment in
 * stable display order; replies are nested under their parent. Each
 * thread exposes scroll-to-highlight (clicking the comment header),
 * inline reply input, resolve/reopen, and delete.
 *
 * Resolved comments are visually de-emphasised (lower contrast, dashed
 * border, "Resolved" pill) but stay visible; clicking "Reopen" inverts
 * the state.
 *
 * The component is provider-driven so DOCX, XLSX, and PPTX editors can
 * all reuse it — pass a `CommentsProvider` and the UI dispatches add /
 * reply / resolve / delete through the adapter's command bus.
 */
export function CommentsSidebar(props: CommentsSidebarProps): ReactNode {
  const threads = resolveThreads(props);
  const author = props.author ?? "You";
  const scrollTo = props.onScrollTo ?? props.provider?.onScrollTo ?? noop;

  const onReply = async (parentId: string, text: string) => {
    if (!props.provider) return;
    await props.provider.reply({ parentId, author, text });
  };
  const onResolve = async (commentId: string, resolved: boolean) => {
    if (!props.provider) return;
    await props.provider.resolve(commentId, resolved);
  };
  const onDelete = async (commentId: string) => {
    if (!props.provider) return;
    await props.provider.delete(commentId);
  };

  return (
    <div className="comments-sidebar flex min-h-0 flex-col gap-2" data-testid="comments-sidebar">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-1.5 text-xs font-medium uppercase tracking-wide text-secondary">
          <MessageCircle size={12} />
          Comments
        </div>
        <span className="rounded-full bg-hover px-2 py-0.5 text-[10px] font-medium text-secondary">
          {threads.length}
        </span>
      </div>
      {threads.length === 0 ? (
        <p className="text-xs text-secondary">
          {props.emptyHint ?? "No comments yet."}
        </p>
      ) : (
        <ul className="flex flex-col gap-2 overflow-y-auto pr-1">
          {threads.map((thread) => (
            <li key={thread.parent.id}>
              <CommentThreadCard
                thread={thread}
                onScrollTo={scrollTo}
                onReply={onReply}
                onResolve={onResolve}
                onDelete={onDelete}
              />
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

function resolveThreads(props: CommentsSidebarProps): ReadonlyArray<CommentThread> {
  if (props.provider) return props.provider.threads();
  return props.threads ?? [];
}

function noop(): void {
  /* default no-op scroll handler */
}

function CommentThreadCard(props: {
  thread: CommentThread;
  onScrollTo: (commentId: string) => void;
  onReply: (parentId: string, text: string) => Promise<void> | void;
  onResolve: (commentId: string, resolved: boolean) => Promise<void> | void;
  onDelete: (commentId: string) => Promise<void> | void;
}): ReactNode {
  const { thread } = props;
  const resolved = !!thread.parent.resolved;
  return (
    <div
      data-testid="comment-thread"
      data-comment-id={thread.parent.id}
      data-resolved={resolved ? "true" : "false"}
      className={cn(
        "rounded-md border bg-surface p-2 text-xs shadow-sm transition-opacity",
        resolved ? "border-dashed border-divider opacity-60" : "border-divider"
      )}
    >
      <CommentRow
        comment={thread.parent}
        onScrollTo={props.onScrollTo}
        actions={
          <>
            <button
              type="button"
              title={resolved ? "Reopen" : "Resolve"}
              aria-label={resolved ? "Reopen comment" : "Resolve comment"}
              onClick={() => void props.onResolve(thread.parent.id, !resolved)}
              className="rounded p-1 text-[var(--success)] hover:bg-[var(--success)]/10"
            >
              <Check size={12} />
            </button>
            <button
              type="button"
              title="Delete"
              aria-label="Delete comment"
              onClick={() => void props.onDelete(thread.parent.id)}
              className="rounded p-1 text-[var(--error)] hover:bg-[var(--error)]/10"
            >
              <Trash2 size={12} />
            </button>
          </>
        }
        badge={
          resolved ? (
            <span className="rounded-full bg-[var(--success)]/10 px-1.5 py-0.5 text-[10px] font-medium text-[var(--success)]">
              Resolved
            </span>
          ) : null
        }
      />
      {thread.replies.length > 0 && (
        <ul className="mt-2 flex flex-col gap-1 border-l border-divider pl-2">
          {thread.replies.map((reply) => (
            <li key={reply.id}>
              <CommentRow
                comment={reply}
                onScrollTo={props.onScrollTo}
                indent
                actions={
                  <button
                    type="button"
                    title="Delete reply"
                    aria-label="Delete reply"
                    onClick={() => void props.onDelete(reply.id)}
                    className="rounded p-1 text-[var(--error)] hover:bg-[var(--error)]/10"
                  >
                    <Trash2 size={12} />
                  </button>
                }
              />
            </li>
          ))}
        </ul>
      )}
      {!resolved && <ReplyInput parentId={thread.parent.id} onSubmit={props.onReply} />}
    </div>
  );
}

function CommentRow(props: {
  comment: CommentBody;
  onScrollTo: (id: string) => void;
  actions?: ReactNode;
  badge?: ReactNode;
  indent?: boolean;
}): ReactNode {
  const text = props.comment.text;
  return (
    <div className={cn("flex items-start gap-2", props.indent && "text-[11px]")}>
      {props.indent && <CornerDownRight size={10} className="mt-1 shrink-0 text-tertiary" />}
      <div className="min-w-0 flex-1">
        <div className="flex items-center gap-1.5">
          <button
            type="button"
            onClick={() => props.onScrollTo(props.comment.id)}
            className="truncate font-medium text-foreground hover:underline"
            title={`Scroll to ${props.comment.author}'s comment`}
          >
            {props.comment.author || "Anonymous"}
          </button>
          {props.badge}
        </div>
        <p className="mt-0.5 whitespace-pre-wrap break-words text-secondary">{text}</p>
      </div>
      {props.actions && <div className="flex shrink-0 items-center gap-0.5">{props.actions}</div>}
    </div>
  );
}

function ReplyInput(props: {
  parentId: string;
  onSubmit: (parentId: string, text: string) => Promise<void> | void;
}): ReactNode {
  const [value, setValue] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async (e: FormEvent) => {
    e.preventDefault();
    if (!value.trim() || busy) return;
    setBusy(true);
    try {
      await props.onSubmit(props.parentId, value.trim());
      setValue("");
    } finally {
      setBusy(false);
    }
  };

  return (
    <form onSubmit={submit} className="mt-2 flex items-center gap-1">
      <input
        type="text"
        value={value}
        onChange={(e) => setValue(e.target.value)}
        placeholder="Reply…"
        aria-label="Reply"
        className="min-w-0 flex-1 rounded-md border border-divider bg-background px-2 py-1 text-xs text-foreground placeholder:text-tertiary focus:border-accent focus:outline-none"
      />
      <Button
        type="submit"
        variant="ghost"
        size="sm"
        disabled={busy || value.trim().length === 0}
        aria-label="Send reply"
      >
        Send
      </Button>
    </form>
  );
}

// Re-export so adapters can compose threads themselves if they don't
// want to hand the provider to the sidebar. Mirrors the helpers
// available at @officeai/comments.
export { groupThreads };
