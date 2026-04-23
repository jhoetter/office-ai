"use client";

/**
 * Sheet view dialog.
 *
 * Mirrors Excel's View tab toggle batch in a single modal so the
 * palette / CLI counterpart of `xlsx:set-sheet-view` has a one-stop
 * surface. Lets the user pick the view mode (Normal / Page Break
 * Preview / Page Layout) plus toggle gridlines, headings, the ruler
 * and right-to-left at once. Submission emits a single
 * `xlsx:set-sheet-view` command.
 */

import { useEffect, useRef, useState } from "react";
import type { ReactNode } from "react";
import { X } from "lucide-react";
import { useFocusTrap } from "@officeai/ui";

export type SheetViewMode = "normal" | "pageBreakPreview" | "pageLayout";

export interface SheetViewSubmit {
  readonly view?: SheetViewMode;
  readonly showGridLines?: boolean;
  readonly showRowColHeaders?: boolean;
  readonly showRuler?: boolean;
  readonly rightToLeft?: boolean;
}

export interface SheetViewDialogProps {
  readonly open: boolean;
  readonly current: {
    readonly view: SheetViewMode;
    readonly showGridLines: boolean;
    readonly showRowColHeaders: boolean;
    readonly showRuler: boolean;
    readonly rightToLeft: boolean;
  };
  readonly onClose: () => void;
  readonly onSubmit: (patch: SheetViewSubmit) => void;
}

export function SheetViewDialog(props: SheetViewDialogProps): ReactNode {
  const { open, current, onClose, onSubmit } = props;
  const [view, setView] = useState<SheetViewMode>(current.view);
  const [showGridLines, setShowGridLines] = useState(current.showGridLines);
  const [showRowColHeaders, setShowRowColHeaders] = useState(current.showRowColHeaders);
  const [showRuler, setShowRuler] = useState(current.showRuler);
  const [rightToLeft, setRightToLeft] = useState(current.rightToLeft);
  const panelRef = useRef<HTMLDivElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onClose });

  useEffect(() => {
    if (!open) return;
    setView(current.view);
    setShowGridLines(current.showGridLines);
    setShowRowColHeaders(current.showRowColHeaders);
    setShowRuler(current.showRuler);
    setRightToLeft(current.rightToLeft);
  }, [open, current]);

  if (!open) return null;

  const handleApply = (): void => {
    // Only emit fields that actually changed so the command produces
    // a clean diff (and doesn't bump revision when the user just
    // opened+OK'd without changing anything).
    const patch: { -readonly [K in keyof SheetViewSubmit]: SheetViewSubmit[K] } = {};
    if (view !== current.view) patch.view = view;
    if (showGridLines !== current.showGridLines) patch.showGridLines = showGridLines;
    if (showRowColHeaders !== current.showRowColHeaders) patch.showRowColHeaders = showRowColHeaders;
    if (showRuler !== current.showRuler) patch.showRuler = showRuler;
    if (rightToLeft !== current.rightToLeft) patch.rightToLeft = rightToLeft;
    onSubmit(patch);
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="xlsx-sheet-view-title"
      className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 px-4 py-6"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        ref={panelRef}
        tabIndex={-1}
        className="flex w-full max-w-sm flex-col rounded-lg border border-divider bg-surface shadow-2xl outline-none"
        onClick={(e) => e.stopPropagation()}
        data-testid="xlsx-sheet-view-dialog"
      >
        <header className="flex items-center justify-between gap-3 border-b border-divider px-5 py-3">
          <h2 id="xlsx-sheet-view-title" className="text-base font-semibold">
            Sheet view
          </h2>
          <button
            type="button"
            aria-label="Close"
            onClick={onClose}
            className="rounded p-1 text-secondary transition-colors hover:bg-hover hover:text-default"
          >
            <X size={16} />
          </button>
        </header>
        <div className="flex flex-col gap-4 px-5 py-4 text-sm">
          <fieldset className="flex flex-col gap-1">
            <legend className="mb-1 text-xs font-medium uppercase text-secondary">View mode</legend>
            {(["normal", "pageBreakPreview", "pageLayout"] as const).map((m) => (
              <label key={m} className="inline-flex items-center gap-2">
                <input
                  type="radio"
                  data-testid={`sheet-view-mode-${m}`}
                  checked={view === m}
                  onChange={() => setView(m)}
                />
                <span>
                  {m === "normal"
                    ? "Normal"
                    : m === "pageBreakPreview"
                      ? "Page Break Preview"
                      : "Page Layout"}
                </span>
              </label>
            ))}
          </fieldset>
          <fieldset className="flex flex-col gap-1">
            <legend className="mb-1 text-xs font-medium uppercase text-secondary">Show</legend>
            <label className="inline-flex items-center gap-2">
              <input
                type="checkbox"
                data-testid="sheet-view-gridlines"
                checked={showGridLines}
                onChange={(e) => setShowGridLines(e.currentTarget.checked)}
              />
              <span>Gridlines</span>
            </label>
            <label className="inline-flex items-center gap-2">
              <input
                type="checkbox"
                data-testid="sheet-view-headings"
                checked={showRowColHeaders}
                onChange={(e) => setShowRowColHeaders(e.currentTarget.checked)}
              />
              <span>Headings (row/column)</span>
            </label>
            <label className="inline-flex items-center gap-2">
              <input
                type="checkbox"
                data-testid="sheet-view-ruler"
                checked={showRuler}
                onChange={(e) => setShowRuler(e.currentTarget.checked)}
              />
              <span>Ruler (Page Layout only)</span>
            </label>
            <label className="inline-flex items-center gap-2">
              <input
                type="checkbox"
                data-testid="sheet-view-rtl"
                checked={rightToLeft}
                onChange={(e) => setRightToLeft(e.currentTarget.checked)}
              />
              <span>Right-to-left</span>
            </label>
          </fieldset>
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
            onClick={handleApply}
            data-testid="sheet-view-apply"
            className="rounded bg-accent px-3 py-1.5 text-sm font-medium text-white hover:opacity-90"
          >
            OK
          </button>
        </footer>
      </div>
    </div>
  );
}
