"use client";

import { useState, type FormEvent, type ReactNode } from "react";
import type { CommentAnchor, CommentsProvider } from "@officeai/comments";
import { Button } from "./button";

export interface CommentComposerProps {
  provider: CommentsProvider;
  /** Author display name to attach to new comments. Defaults to "You". */
  author?: string;
  /** Where the new comment should be anchored. */
  anchor: CommentAnchor;
  /** Optional callback fired after the new comment is dispatched. */
  onSubmitted?: (commentId: string) => void;
  /** Placeholder text. Defaults to "Add a comment…". */
  placeholder?: string;
  /** Submit button label. Defaults to "Comment". */
  submitLabel?: string;
}

/**
 * Inline textarea + submit button for adding a fresh comment at a given
 * anchor. Adapter-agnostic: the provider decides how the new comment
 * gets persisted (DOCX adds a `<w:comment>`, XLSX writes a threaded
 * comment, PPTX inserts a slide-level pin). Multi-line input is
 * supported; Enter inserts a newline, Cmd/Ctrl+Enter submits.
 */
export function CommentComposer(props: CommentComposerProps): ReactNode {
  const [text, setText] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async (e?: FormEvent) => {
    if (e) e.preventDefault();
    const trimmed = text.trim();
    if (!trimmed || busy) return;
    setBusy(true);
    try {
      const id = await props.provider.add({
        author: props.author ?? "You",
        text: trimmed,
        anchor: props.anchor,
      });
      setText("");
      props.onSubmitted?.(id);
    } finally {
      setBusy(false);
    }
  };

  return (
    <form
      onSubmit={submit}
      className="flex flex-col gap-1.5 rounded-md border border-divider bg-surface p-2"
      data-testid="comment-composer"
    >
      <textarea
        value={text}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if ((e.metaKey || e.ctrlKey) && e.key === "Enter") {
            e.preventDefault();
            void submit();
          }
        }}
        placeholder={props.placeholder ?? "Add a comment…"}
        aria-label="Comment text"
        className="min-h-[44px] w-full resize-y rounded border border-divider bg-background p-2 text-xs text-foreground placeholder:text-tertiary focus:border-accent focus:outline-none"
      />
      <div className="flex items-center justify-end gap-1">
        <Button
          type="submit"
          variant="ghost"
          size="sm"
          disabled={busy || text.trim().length === 0}
        >
          {props.submitLabel ?? "Comment"}
        </Button>
      </div>
    </form>
  );
}
