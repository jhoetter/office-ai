"use client";

import { useEffect, useRef, useState, type ReactNode } from "react";
import { MessageSquarePlus, X } from "lucide-react";
import { Button } from "@officeai/ui";

/**
 * Comment composer popover.
 *
 * Pure human-driven authoring surface: the user types the comment they
 * actually want and clicks Add. The earlier "Draft with AI" button has
 * been removed — agent-driven comment authoring now flows through the
 * `office-agent` CLI (`oa docx comment`) rather than through the
 * editor UI.
 *
 * The component is purely presentational: positioning + selection
 * snippet come from the parent (`DocxEditor`), which owns the actual
 * command dispatch so the comment lands through the same
 * `agent.applyCommand` funnel as every other mutation.
 */

export interface CommentComposerProps {
  /** Snippet of selected text rendered as a quote chip. Empty string is allowed. */
  selectionText: string;
  /** Initial textarea value. Defaults to empty. */
  initialText?: string;
  /** Called with the final comment body when the user clicks Add. */
  onSubmit: (text: string) => void;
  /** Called when the user dismisses the popover. */
  onCancel: () => void;
  /** Anchor coords (page-relative). The popover positions below it. */
  anchor: { left: number; top: number; bottom: number } | null;
}

export function CommentComposer(props: CommentComposerProps): ReactNode {
  const [text, setText] = useState(props.initialText ?? "");
  const taRef = useRef<HTMLTextAreaElement | null>(null);

  useEffect(() => {
    taRef.current?.focus();
  }, []);

  useEffect(() => {
    function onKey(e: KeyboardEvent): void {
      if (e.key === "Escape") {
        e.preventDefault();
        props.onCancel();
      }
    }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [props]);

  const submit = (): void => {
    const trimmed = text.trim();
    if (!trimmed) return;
    props.onSubmit(trimmed);
  };

  // Anchor: `coordsAtPos` from PM returns viewport-relative coordinates,
  // so we use position:fixed for the anchored case as well. Falling back
  // to a centred dialog when no anchor is available means the popover
  // is always reachable even from keyboard-only flows.
  const style: React.CSSProperties = props.anchor
    ? {
        position: "fixed",
        left: Math.max(
          8,
          Math.min(props.anchor.left, (typeof window !== "undefined" ? window.innerWidth : 1024) - 340)
        ),
        top: props.anchor.bottom + 6,
      }
    : { position: "fixed", left: "50%", top: "30%", transform: "translate(-50%,-30%)" };

  return (
    <div
      data-testid="comment-composer"
      role="dialog"
      aria-label="Add comment"
      className="z-50 flex w-[320px] flex-col gap-2 rounded-md border border-divider bg-surface p-3 shadow-lg"
      style={style}
    >
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-1.5 text-xs font-medium uppercase tracking-wide text-secondary">
          <MessageSquarePlus size={12} /> Comment
        </div>
        <button
          type="button"
          aria-label="Cancel"
          onClick={props.onCancel}
          className="rounded p-1 text-secondary hover:bg-divider/50"
        >
          <X size={12} />
        </button>
      </div>

      {props.selectionText && (
        <blockquote
          className="max-h-16 overflow-hidden text-ellipsis rounded border-l-2 border-[var(--ai-violet)] bg-background/60 px-2 py-1 text-[11px] italic text-secondary"
          title={props.selectionText}
        >
          {truncate(props.selectionText, 160)}
        </blockquote>
      )}

      <textarea
        ref={taRef}
        data-testid="comment-composer-text"
        value={text}
        onChange={(e) => setText(e.target.value)}
        rows={3}
        placeholder="Type your comment…"
        className="w-full resize-none rounded-md border border-divider bg-background px-2 py-1.5 text-xs text-foreground placeholder:text-tertiary focus:border-[var(--ai-violet)] focus:outline-none"
        onKeyDown={(e) => {
          if ((e.metaKey || e.ctrlKey) && e.key === "Enter") {
            e.preventDefault();
            submit();
          }
        }}
      />

      <div className="flex items-center justify-end gap-1.5">
        <Button variant="ghost" size="sm" onClick={props.onCancel}>
          Cancel
        </Button>
        <Button
          variant="accent"
          size="sm"
          onClick={submit}
          disabled={text.trim().length === 0}
          data-testid="comment-composer-submit"
        >
          Add comment
        </Button>
      </div>
    </div>
  );
}

function truncate(s: string, n: number): string {
  if (s.length <= n) return s;
  return `${s.slice(0, n - 1).trimEnd()}…`;
}
