"use client";

import { useEffect, useState, type ReactNode } from "react";
import { X } from "lucide-react";
import { cn } from "@officeai/ui";

/**
 * C7 — Paste Special dialog.
 *
 * Opens on Cmd+Shift+V (Excel parity). Lets the user pick what to
 * paste from the system clipboard (or in-app copy):
 *
 *   - All           : values + formulas + formats + merges
 *   - Values        : computed values only — formulas collapse
 *   - Formulas      : formulas (relative-shifted) without source styles
 *   - Formats       : per-cell style only, never overwrites values
 *
 * Plus a Transpose toggle that flips rows/columns.
 *
 * The dialog is intentionally thin: it owns no clipboard state. The
 * parent reads the clipboard at confirm-time and dispatches
 * `xlsx:paste-range` with the chosen `{ mode, transpose }` pair.
 */

export type PasteSpecialMode = "all" | "values" | "formulas" | "formats";

export interface PasteSpecialOptions {
  readonly mode: PasteSpecialMode;
  readonly transpose: boolean;
}

export interface PasteSpecialDialogProps {
  readonly open: boolean;
  readonly onClose: () => void;
  readonly onConfirm: (opts: PasteSpecialOptions) => void;
}

interface ModeOption {
  readonly id: PasteSpecialMode;
  readonly label: string;
  readonly hint: string;
}

const MODE_OPTIONS: ReadonlyArray<ModeOption> = [
  {
    id: "all",
    label: "All",
    hint: "Values, formulas, formats, and merged ranges.",
  },
  {
    id: "values",
    label: "Values",
    hint: "Numbers and text only — formulas collapse to their result.",
  },
  {
    id: "formulas",
    label: "Formulas",
    hint: "Formulas (with relative refs adjusted) without copying styles.",
  },
  {
    id: "formats",
    label: "Formats",
    hint: "Cell formatting only — does not overwrite the destination values.",
  },
];

export function PasteSpecialDialog(props: PasteSpecialDialogProps): ReactNode {
  const { open, onClose, onConfirm } = props;
  const [mode, setMode] = useState<PasteSpecialMode>("all");
  const [transpose, setTranspose] = useState<boolean>(false);

  // Reset to defaults each time the dialog opens so the user gets a
  // predictable starting state. (Excel remembers the last choice; we
  // can wire that to localStorage later if it becomes a habit.)
  useEffect(() => {
    if (!open) return;
    setMode("all");
    setTranspose(false);
  }, [open]);

  // Esc / Enter shortcuts. Enter commits the current selection so a
  // pure keyboard flow is "Cmd+Shift+V → arrow → Enter".
  useEffect(() => {
    if (!open) return;
    const handler = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        onClose();
        return;
      }
      if (e.key === "Enter") {
        e.preventDefault();
        onConfirm({ mode, transpose });
        return;
      }
    };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, [open, onClose, onConfirm, mode, transpose]);

  if (!open) return null;

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Paste Special"
      data-testid="paste-special-dialog"
      className="fixed inset-0 z-50 flex items-center justify-center bg-foreground/30 p-4"
      onMouseDown={(e) => {
        // Click on the backdrop dismisses; click inside the panel
        // doesn't bubble because the panel stops propagation below.
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        className="w-[420px] max-w-full rounded-lg border border-divider bg-surface shadow-xl"
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between border-b border-divider px-4 py-2.5">
          <h2 className="text-sm font-semibold text-foreground">Paste Special</h2>
          <button
            type="button"
            aria-label="Close"
            data-testid="paste-special-close"
            onClick={onClose}
            className="inline-flex h-7 w-7 items-center justify-center rounded text-secondary hover:bg-hover"
          >
            <X size={14} />
          </button>
        </div>
        <div className="flex flex-col gap-1 p-3">
          {MODE_OPTIONS.map((opt) => (
            <label
              key={opt.id}
              data-testid={`paste-special-${opt.id}`}
              className={cn(
                "flex cursor-pointer items-start gap-3 rounded-md border border-transparent px-3 py-2 hover:bg-hover",
                mode === opt.id && "border-divider bg-[var(--ai-violet-light)]"
              )}
            >
              <input
                type="radio"
                name="paste-special-mode"
                value={opt.id}
                checked={mode === opt.id}
                onChange={() => setMode(opt.id)}
                className="mt-0.5 accent-[var(--ai-violet)]"
              />
              <div className="flex flex-col gap-0.5">
                <span className="text-xs font-medium text-foreground">{opt.label}</span>
                <span className="text-[11px] text-tertiary">{opt.hint}</span>
              </div>
            </label>
          ))}
        </div>
        <div className="border-t border-divider px-3 py-2">
          <label
            data-testid="paste-special-transpose"
            className="flex cursor-pointer items-center gap-2 px-1 py-1"
          >
            <input
              type="checkbox"
              checked={transpose}
              onChange={(e) => setTranspose(e.target.checked)}
              className="accent-[var(--ai-violet)]"
            />
            <span className="text-xs text-foreground">Transpose (rows ↔ columns)</span>
          </label>
        </div>
        <div className="flex items-center justify-end gap-2 border-t border-divider px-3 py-2.5">
          <button
            type="button"
            onClick={onClose}
            data-testid="paste-special-cancel"
            className="inline-flex h-7 items-center rounded px-3 text-xs text-foreground hover:bg-hover"
          >
            Cancel
          </button>
          <button
            type="button"
            data-testid="paste-special-confirm"
            onClick={() => onConfirm({ mode, transpose })}
            className="inline-flex h-7 items-center rounded bg-[var(--ai-violet)] px-3 text-xs font-medium text-white hover:opacity-90"
          >
            Paste
          </button>
        </div>
      </div>
    </div>
  );
}
