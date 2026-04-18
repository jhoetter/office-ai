"use client";

import { useEffect, useRef, useState, type ReactNode } from "react";
import { Loader2, MessageSquarePlus, Sparkles, X } from "lucide-react";
import { Button } from "@officeai/ui";

/**
 * Comment composer popover (P2.5 / W24).
 *
 * Replaces the old hard-coded `text: "Looks good?"` recipe with a
 * proper composer surface. The user types the comment they actually
 * want; an optional "Draft with AI" button asks the LLM bridge to draft
 * a constructive comment about the highlighted snippet (the bridge is
 * passed the same selection context as `dispatchToLlm`, so when the
 * server has no API key the offline fallback simply mirrors the
 * snippet back as a placeholder draft).
 *
 * The component is purely presentational: positioning + selection
 * snippet + busy-state come from the parent (`DocxEditor`). The parent
 * owns the actual command dispatch so it can route the comment through
 * the same `agent.applyCommand` funnel as everything else.
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
  /**
   * Optional AI-draft helper. When provided, an "Ask AI" button appears
   * that calls this with the current draft and the selection snippet,
   * and the returned text replaces the textarea contents.
   */
  onDraftWithAi?: (currentDraft: string, selectionText: string) => Promise<string>;
  /** Anchor coords (page-relative). The popover positions below it. */
  anchor: { left: number; top: number; bottom: number } | null;
}

export function CommentComposer(props: CommentComposerProps): ReactNode {
  const [text, setText] = useState(props.initialText ?? "");
  const [busy, setBusy] = useState(false);
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

  const draft = async (): Promise<void> => {
    if (!props.onDraftWithAi) return;
    setBusy(true);
    try {
      const next = await props.onDraftWithAi(text, props.selectionText);
      if (typeof next === "string" && next.length > 0) setText(next);
    } finally {
      setBusy(false);
    }
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

      <div className="flex items-center justify-between gap-2">
        {props.onDraftWithAi ? (
          <Button
            variant="ghost"
            size="sm"
            onClick={() => void draft()}
            disabled={busy}
            className="text-[var(--ai-violet)] hover:bg-[var(--ai-violet)]/10"
            data-testid="comment-composer-ai"
          >
            {busy ? <Loader2 className="animate-spin" size={12} /> : <Sparkles size={12} />}
            Draft with AI
          </Button>
        ) : (
          <span />
        )}
        <div className="flex items-center gap-1.5">
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
    </div>
  );
}

function truncate(s: string, n: number): string {
  if (s.length <= n) return s;
  return `${s.slice(0, n - 1).trimEnd()}…`;
}
