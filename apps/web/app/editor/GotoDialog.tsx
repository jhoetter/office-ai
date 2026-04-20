"use client";

import { useEffect, useRef, useState, type ReactNode } from "react";
import { cn, useFocusTrap } from "@officeai/ui";
import { useTranslator } from "@/lib/i18n";

export interface GotoDialogProps {
  readonly open: boolean;
  readonly currentPage: number;
  readonly totalPages: number;
  readonly onSubmit: (page: number) => void;
  readonly onClose: () => void;
}

/**
 * B10 — minimal "Go to page" dialog. Word's `Mod+G` lands the user
 * here: a tiny modal with a single number input that snaps the
 * caret onto the requested page (clamped to `[1, totalPages]`).
 *
 * Kept intentionally lean — Word's "Find and Replace › Go To" tab
 * has options for sections, lines, comments, etc.; we don't need
 * any of that today and can grow the surface when a user actually
 * asks for it. The current page is pre-filled so a stray Enter is
 * a harmless no-op.
 */
export function GotoDialog({ open, currentPage, totalPages, onSubmit, onClose }: GotoDialogProps): ReactNode {
  const { t } = useTranslator();
  const [draft, setDraft] = useState(String(currentPage));
  const inputRef = useRef<HTMLInputElement | null>(null);
  const panelRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    setDraft(String(currentPage));
    const handle = window.setTimeout(() => {
      inputRef.current?.select();
    }, 0);
    return () => window.clearTimeout(handle);
  }, [open, currentPage]);

  useFocusTrap(panelRef, {
    enabled: open,
    initialFocusRef: inputRef,
    onEscape: onClose,
  });

  if (!open) return null;

  const submit = () => {
    const n = Number.parseInt(draft, 10);
    if (Number.isFinite(n)) {
      const clamped = Math.max(1, Math.min(totalPages || 1, n));
      onSubmit(clamped);
    }
    onClose();
  };

  return (
    <div
      role="presentation"
      className="fixed inset-0 z-40 flex items-start justify-center bg-black/30 p-4 pt-[20vh]"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
      data-testid="goto-dialog"
    >
      <div
        ref={panelRef}
        tabIndex={-1}
        role="dialog"
        aria-modal="true"
        aria-label={t("docx.goto.title")}
        className={cn(
          "w-full max-w-sm rounded-lg border border-divider bg-surface p-4 shadow-lg outline-none"
        )}
      >
        <h2 className="text-sm font-semibold text-foreground">{t("docx.goto.title")}</h2>
        <p className="mt-1 text-xs text-secondary">{t("docx.goto.hint", { total: totalPages || 1 })}</p>
        <div className="mt-3 flex items-center gap-2">
          <input
            ref={inputRef}
            type="number"
            min={1}
            max={totalPages || 1}
            value={draft}
            onChange={(e) => setDraft(e.currentTarget.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                e.preventDefault();
                submit();
              }
            }}
            className="w-24 rounded-md border border-divider bg-background px-2 py-1 text-sm text-foreground focus:border-[var(--accent)] focus:outline-none"
            data-testid="goto-dialog-input"
            aria-label={t("docx.goto.pageNumber")}
          />
          <span className="text-xs text-secondary">{t("docx.goto.of", { total: totalPages || 1 })}</span>
        </div>
        <div className="mt-4 flex justify-end gap-2">
          <button
            type="button"
            onClick={onClose}
            className="rounded-md border border-divider bg-surface px-3 py-1 text-xs font-medium text-foreground hover:bg-hover"
            data-testid="goto-dialog-cancel"
          >
            {t("common.cancel")}
          </button>
          <button
            type="button"
            onClick={submit}
            className="rounded-md bg-[var(--accent)] px-3 py-1 text-xs font-medium text-white hover:opacity-90"
            data-testid="goto-dialog-submit"
          >
            {t("docx.goto.go")}
          </button>
        </div>
      </div>
    </div>
  );
}
