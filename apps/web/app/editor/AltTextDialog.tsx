"use client";

import { useEffect, useRef, useState } from "react";
import { X } from "lucide-react";
import { useFocusTrap } from "@officeai/ui";

/**
 * B6 — alt-text editor.
 *
 * Word's "Edit alt text" pane reduced to its essence: a short
 * description (one or two sentences), a "mark as decorative" toggle
 * and the standard Cancel / OK affordances. We focus the textarea on
 * open so power users can tab in from the contextual toolbar and
 * type immediately. Submitting commits a single
 * `docx:set-image-properties` with `altText: …` (or `null` when
 * decorative).
 */

export interface AltTextDialogProps {
  readonly open: boolean;
  readonly initial: string;
  readonly imageId: string | null;
  readonly onClose: () => void;
  readonly onSubmit: (imageId: string, altText: string | null) => void;
}

export function AltTextDialog(props: AltTextDialogProps) {
  const { open, initial, imageId, onClose, onSubmit } = props;
  const [draft, setDraft] = useState<string>(initial);
  const [decorative, setDecorative] = useState<boolean>(false);
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);
  const panelRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    setDraft(initial);
    setDecorative(initial.length === 0);
  }, [open, initial]);

  useFocusTrap(panelRef, {
    enabled: open,
    initialFocusRef: textareaRef,
    onEscape: onClose,
  });

  if (!open || !imageId) return null;

  const submit = () => {
    onSubmit(imageId, decorative ? null : draft);
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="alt-text-title"
      className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 px-4 py-6"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        ref={panelRef}
        tabIndex={-1}
        className="flex w-full max-w-md flex-col rounded-lg border border-divider bg-surface shadow-2xl outline-none"
        onClick={(e) => e.stopPropagation()}
      >
        <header className="flex items-center justify-between gap-3 border-b border-divider px-5 py-3">
          <div>
            <h2 id="alt-text-title" className="text-base font-semibold">
              Edit alt text
            </h2>
            <p className="text-xs text-secondary">
              Describe the image so screen readers can convey its meaning.
            </p>
          </div>
          <button
            type="button"
            aria-label="Close"
            onClick={onClose}
            className="rounded p-1 text-secondary transition-colors hover:bg-hover hover:text-default"
          >
            <X size={16} />
          </button>
        </header>

        <div className="flex flex-col gap-3 px-5 py-4 text-sm">
          <label className="flex flex-col gap-1.5">
            <span className="text-xs font-medium uppercase tracking-wide text-secondary">Description</span>
            <textarea
              ref={textareaRef}
              value={draft}
              onChange={(e) => {
                setDraft(e.target.value);
                if (e.target.value.length > 0) setDecorative(false);
              }}
              rows={4}
              className="resize-y rounded border border-divider bg-background px-2 py-1.5"
              placeholder="e.g. Bar chart comparing Q1–Q4 revenue."
              data-testid="alt-text-input"
              disabled={decorative}
            />
          </label>
          <label className="flex items-center gap-2">
            <input
              type="checkbox"
              checked={decorative}
              onChange={(e) => {
                setDecorative(e.target.checked);
                if (e.target.checked) setDraft("");
              }}
            />
            <span>Mark as decorative (skipped by screen readers)</span>
          </label>
        </div>

        <footer className="flex items-center justify-end gap-2 border-t border-divider px-5 py-3">
          <button
            type="button"
            onClick={onClose}
            className="rounded border border-divider bg-background px-3 py-1.5 text-sm hover:bg-hover"
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={submit}
            className="rounded border border-transparent bg-accent px-3 py-1.5 text-sm font-medium text-white hover:opacity-90"
            data-testid="alt-text-submit"
          >
            Save
          </button>
        </footer>
      </div>
    </div>
  );
}
