"use client";

import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { Trash2, Pencil, Check, X } from "@officeai/ui/sonaloop-icons";
import { cn } from "@officeai/ui";
import type { DocxAgent, DocxSnapshot, Footnote } from "@officeai/docx";

/**
 * 9b — In-place footnote panel rendered in the right rail (Footnotes
 * tab) for DOCX. Lists every user-authored footnote in document order
 * and exposes the two minimum operations users expect from Word's
 * Review › Footnotes pane:
 *
 *   - inline-edit a footnote body (dispatches `docx:set-footnote-body`)
 *   - delete a footnote AND every reference to it in the body
 *     (dispatches `docx:delete-footnote`)
 *
 * The standard separator (id = -1) and continuation (id = 0) auto-
 * notes Word inserts on every footnote-bearing document are filtered
 * out — they're not user content and exposing them would only confuse.
 *
 * Body editing is intentionally limited to plain text. Footnotes
 * usually carry a single short paragraph; richer formatting is a
 * future extension that would require a mini ProseMirror surface.
 * The plain-text editor preserves paragraph splits on newline so
 * multi-line footnotes can still round-trip without losing structure.
 */
export interface FootnotesPanelProps {
  readonly agent: DocxAgent | null;
  readonly snapshot: DocxSnapshot | null;
  /**
   * Optional click-to-scroll hook. Receives the footnote id; the
   * caller resolves the matching `<w:footnoteReference>` in the body
   * and brings it into view. We don't reach into ProseMirror from
   * the rail directly — the editor owns the view.
   */
  readonly onScrollToReference?: (footnoteId: number) => void;
}

export function FootnotesPanel(props: FootnotesPanelProps): ReactNode {
  const footnotes = useUserFootnotes(props.snapshot);

  if (footnotes.length === 0) {
    return (
      <div className="p-4 text-sm text-secondary" data-testid="footnotes-panel-empty">
        No footnotes yet. Insert one via References → Footnote.
      </div>
    );
  }

  return (
    <ul className="flex min-h-0 flex-1 flex-col gap-1 overflow-y-auto p-2" data-testid="footnotes-panel">
      {footnotes.map((fn) => (
        <FootnoteRow
          key={fn.id}
          footnote={fn}
          agent={props.agent}
          onScrollToReference={props.onScrollToReference}
        />
      ))}
    </ul>
  );
}

function FootnoteRow({
  footnote,
  agent,
  onScrollToReference,
}: {
  readonly footnote: Footnote;
  readonly agent: DocxAgent | null;
  readonly onScrollToReference?: (footnoteId: number) => void;
}): ReactNode {
  const text = useMemo(() => footnoteToPlainText(footnote), [footnote]);
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(text);
  const [saving, setSaving] = useState(false);
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);

  // When the snapshot updates externally (someone else edited the
  // footnote, undo, etc.) re-sync the draft so we don't trap a stale
  // value behind the user's pencil click.
  useEffect(() => {
    if (!editing) setDraft(text);
  }, [text, editing]);

  const enterEdit = () => {
    setDraft(text);
    setEditing(true);
    // Focus on next tick so the textarea exists.
    requestAnimationFrame(() => {
      textareaRef.current?.focus();
      textareaRef.current?.select();
    });
  };

  const cancelEdit = () => {
    setDraft(text);
    setEditing(false);
  };

  const commitEdit = async () => {
    if (!agent) return;
    if (draft === text) {
      setEditing(false);
      return;
    }
    setSaving(true);
    try {
      await agent.applyCommand({
        type: "docx:set-footnote-body",
        payload: {
          footnoteId: footnote.id,
          body: plainTextToBlocks(draft),
        },
        source: "human",
      });
      setEditing(false);
    } finally {
      setSaving(false);
    }
  };

  const removeFootnote = async () => {
    if (!agent) return;
    // Footnote deletion also strips every in-text reference; that's a
    // structural change, so we confirm even though there's an undo.
    const ok = window.confirm(
      `Delete footnote ${footnote.id}? Every reference in the body will be removed too.`
    );
    if (!ok) return;
    await agent.applyCommand({
      type: "docx:delete-footnote",
      payload: { footnoteId: footnote.id },
      source: "human",
    });
  };

  return (
    <li className="rounded border border-divider bg-surface p-2" data-testid={`footnote-row-${footnote.id}`}>
      <div className="flex items-start gap-2">
        <button
          type="button"
          onClick={() => onScrollToReference?.(footnote.id)}
          disabled={!onScrollToReference}
          title="Scroll to reference"
          className={cn(
            "shrink-0 rounded bg-hover px-1.5 py-0.5 text-[11px] font-semibold text-foreground hover:bg-divider",
            !onScrollToReference && "cursor-default opacity-60 hover:bg-hover"
          )}
          data-testid={`footnote-jump-${footnote.id}`}
        >
          [{footnote.id}]
        </button>
        <div className="min-w-0 flex-1">
          {editing ? (
            <textarea
              ref={textareaRef}
              value={draft}
              onChange={(e) => setDraft(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Escape") {
                  e.preventDefault();
                  cancelEdit();
                } else if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) {
                  e.preventDefault();
                  void commitEdit();
                }
              }}
              rows={Math.max(2, draft.split("\n").length)}
              className="w-full resize-y rounded border border-divider bg-background px-2 py-1 text-xs text-foreground focus:border-[var(--accent)] focus:outline-none"
              disabled={saving}
              data-testid={`footnote-editor-${footnote.id}`}
            />
          ) : (
            <p
              className="whitespace-pre-wrap break-words text-xs text-foreground"
              data-testid={`footnote-body-${footnote.id}`}
            >
              {text || <span className="text-secondary italic">(empty)</span>}
            </p>
          )}
        </div>
      </div>
      <div className="mt-1 flex justify-end gap-1">
        {editing ? (
          <>
            <button
              type="button"
              onClick={cancelEdit}
              disabled={saving}
              title="Cancel"
              className="inline-flex items-center gap-1 rounded px-1.5 py-0.5 text-[11px] text-secondary hover:bg-hover hover:text-foreground"
              data-testid={`footnote-cancel-${footnote.id}`}
            >
              <X size={11} /> Cancel
            </button>
            <button
              type="button"
              onClick={() => void commitEdit()}
              disabled={saving}
              title="Save (Cmd+Enter)"
              className="inline-flex items-center gap-1 rounded bg-[var(--accent)] px-1.5 py-0.5 text-[11px] font-medium text-white disabled:opacity-50"
              data-testid={`footnote-save-${footnote.id}`}
            >
              <Check size={11} /> Save
            </button>
          </>
        ) : (
          <>
            <button
              type="button"
              onClick={enterEdit}
              disabled={!agent}
              title="Edit"
              className="inline-flex items-center gap-1 rounded px-1.5 py-0.5 text-[11px] text-secondary hover:bg-hover hover:text-foreground disabled:opacity-50"
              data-testid={`footnote-edit-${footnote.id}`}
            >
              <Pencil size={11} /> Edit
            </button>
            <button
              type="button"
              onClick={() => void removeFootnote()}
              disabled={!agent}
              title="Delete"
              className="inline-flex items-center gap-1 rounded px-1.5 py-0.5 text-[11px] text-secondary hover:bg-hover hover:text-red-500 disabled:opacity-50"
              data-testid={`footnote-delete-${footnote.id}`}
            >
              <Trash2 size={11} /> Delete
            </button>
          </>
        )}
      </div>
    </li>
  );
}

/**
 * Filter the footnotes part down to user-authored notes — strips the
 * standard `separator` (id = -1) and `continuationSeparator` (id = 0)
 * Word inserts on every footnote-bearing document. Returns a stable
 * memoised list keyed by snapshot revision so the panel doesn't
 * re-render on unrelated body edits.
 */
export function useUserFootnotes(snapshot: DocxSnapshot | null): ReadonlyArray<Footnote> {
  return useMemo(() => {
    const part = snapshot?.root.footnotesPart;
    if (!part) return [];
    return part.footnotes.filter((f) => f.id > 0);
  }, [snapshot?.root.footnotesPart, snapshot?.revision]);
}

/**
 * Walk a footnote's typed `body` and concatenate the visible text
 * across paragraphs (with `\n` separators). Skips opaque inlines and
 * non-text run children — we never want hidden formatting markers
 * leaking into the editor textarea.
 */
function footnoteToPlainText(footnote: Footnote): string {
  const lines: string[] = [];
  for (const block of footnote.body) {
    if (block.kind !== "paragraph") continue;
    let line = "";
    for (const inline of block.children) {
      if (inline.kind !== "run") continue;
      for (const child of inline.children) {
        if (child.kind === "text") line += child.text;
      }
    }
    lines.push(line);
  }
  return lines.join("\n");
}

/**
 * Convert plain-text input back into the typed BlockNode array
 * `docx:set-footnote-body` expects. Splits on `\n` so multi-line
 * footnotes round-trip; every line becomes its own paragraph with
 * a single text run. The runtime mints fresh ids — the snapshot
 * evolution wraps each block with `_id` machinery, so we hand it
 * the lightest possible structure and let it lay down identifiers.
 *
 * We bake the standard "FootnoteText" paragraph style into each
 * paragraph so newly authored footnotes pick up Word's footnote
 * formatting (small font, hanging indent) instead of inheriting
 * Normal style.
 */
function plainTextToBlocks(text: string): import("@officeai/docx").BlockNode[] {
  const lines = text.length === 0 ? [""] : text.split("\n");
  return lines.map((line) => ({
    kind: "paragraph",
    id: "" as never,
    properties: { styleId: "FootnoteText" },
    children: [
      {
        kind: "run",
        id: "" as never,
        properties: {},
        children: [{ kind: "text", id: "" as never, text: line, xmlSpacePreserve: line.length === 0 }],
      },
    ],
  })) as unknown as import("@officeai/docx").BlockNode[];
}
