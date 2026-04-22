"use client";

/**
 * Excel-style Zoom dialog.
 *
 * Mirrors View → Zoom → Zoom… in Excel. Lets the user pick from
 * standard preset percentages (200/100/75/50/25), enter a custom
 * value (10–400) or "Fit selection". Submission dispatches
 * `xlsx:set-sheet-view` with `zoomScale`.
 */

import { useEffect, useRef, useState } from "react";
import type { ReactNode } from "react";
import { X } from "lucide-react";
import { useFocusTrap } from "@officeai/ui";

export interface ZoomDialogProps {
  readonly open: boolean;
  readonly currentZoom: number;
  readonly onClose: () => void;
  readonly onSubmit: (zoom: number) => void;
}

const PRESETS = [200, 100, 75, 50, 25] as const;

export function ZoomDialog(props: ZoomDialogProps): ReactNode {
  const { open, currentZoom, onClose, onSubmit } = props;
  const [value, setValue] = useState<number>(currentZoom);
  const panelRef = useRef<HTMLDivElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onClose });

  useEffect(() => {
    if (open) setValue(currentZoom);
  }, [open, currentZoom]);

  if (!open) return null;

  const handleApply = (): void => {
    const clamped = Math.max(10, Math.min(400, Math.round(value)));
    onSubmit(clamped);
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="xlsx-zoom-title"
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
      >
        <header className="flex items-center justify-between gap-3 border-b border-divider px-5 py-3">
          <h2 id="xlsx-zoom-title" className="text-base font-semibold">
            Zoom
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
        <div className="flex flex-col gap-3 px-5 py-4 text-sm">
          <div className="flex flex-col gap-1">
            {PRESETS.map((p) => (
              <label key={p} className="flex items-center gap-2">
                <input
                  type="radio"
                  data-testid={`zoom-preset-${p}`}
                  checked={value === p}
                  onChange={() => setValue(p)}
                />
                <span>{p}%</span>
              </label>
            ))}
          </div>
          <div className="flex items-center gap-2">
            <input
              type="radio"
              data-testid="zoom-custom-radio"
              checked={!PRESETS.includes(value as (typeof PRESETS)[number])}
              onChange={() => {}}
              readOnly
            />
            <label className="flex items-center gap-2">
              <span>Custom:</span>
              <input
                type="number"
                min={10}
                max={400}
                value={value}
                onChange={(e) => setValue(Number(e.currentTarget.value) || 100)}
                className="w-20 rounded border border-divider bg-background px-2 py-1 tabular-nums"
                data-testid="zoom-custom-input"
              />
              <span>%</span>
            </label>
          </div>
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
            data-testid="zoom-apply"
            className="rounded bg-accent px-3 py-1.5 text-sm font-medium text-white hover:opacity-90"
          >
            OK
          </button>
        </footer>
      </div>
    </div>
  );
}
