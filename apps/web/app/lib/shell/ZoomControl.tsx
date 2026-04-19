"use client";

import type { ReactElement } from "react";

/**
 * Shared zoom control for the editor status bar.
 *
 * All three products (DOCX, XLSX, PPTX) park zoom in the same place
 * — bottom-right of the status bar — so users get the same
 * affordance no matter which file type they have open. The control
 * itself stays presentational; each product owns its own zoom state
 * and clamping, which it passes in via `value` + `onChange`.
 *
 * Layout matches Office's status bar: minus button, percent label
 * (click to reset to 100 %), plus button. Optional reset button is
 * folded into the percent label to keep the footprint small.
 */
export interface ZoomControlProps {
  readonly value: number;
  readonly minZoom?: number;
  readonly maxZoom?: number;
  readonly step?: number;
  readonly onChange: (next: number) => void;
  readonly disabled?: boolean;
}

export function ZoomControl(props: ZoomControlProps): ReactElement {
  const { value, minZoom = 0.5, maxZoom = 2, step = 0.1, onChange, disabled = false } = props;

  const clamp = (n: number): number => Math.min(maxZoom, Math.max(minZoom, n));
  const pct = Math.round(value * 100);

  return (
    <div className="flex items-center gap-1 text-xs text-secondary">
      <button
        type="button"
        onClick={() => onChange(clamp(Math.round((value - step) * 100) / 100))}
        disabled={disabled || value <= minZoom + 1e-6}
        aria-label="Zoom out"
        className="rounded border border-divider px-1.5 py-0.5 hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
        data-testid="zoom-out"
      >
        −
      </button>
      <button
        type="button"
        onClick={() => onChange(1)}
        disabled={disabled}
        title="Reset zoom to 100%"
        className="min-w-[44px] rounded px-1.5 py-0.5 text-center tabular-nums hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
        data-testid="zoom-percent"
      >
        {pct}%
      </button>
      <button
        type="button"
        onClick={() => onChange(clamp(Math.round((value + step) * 100) / 100))}
        disabled={disabled || value >= maxZoom - 1e-6}
        aria-label="Zoom in"
        className="rounded border border-divider px-1.5 py-0.5 hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
        data-testid="zoom-in"
      >
        +
      </button>
    </div>
  );
}
