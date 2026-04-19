"use client";

import type { ReactNode } from "react";
import { cn } from "@officeai/ui";
import type { ProductAdapter, SaveState } from "./types";

export interface EditorStatusBarProps {
  readonly adapter: ProductAdapter;
  /** Slot for product-specific left content — page n of m, sheet
   * tabs strip, slide n of m. */
  readonly leftSlot?: ReactNode;
  /** Slot for product-specific right content — zoom controls, mode
   * pill, language. */
  readonly rightSlot?: ReactNode;
}

/**
 * The single status bar used by all three editors.
 *
 * Layout: left slot · selection summary (centred) · right slot.
 * The selection summary is a live region — screen readers announce
 * it when it changes. Aggregates render as inline `Sum: 12` chips
 * (XLSX); the simple text variant renders as plain text (DOCX/PPTX).
 */
export function EditorStatusBar({ adapter, leftSlot, rightSlot }: EditorStatusBarProps): ReactNode {
  const summary = adapter.selectionSummary;
  return (
    <footer
      className="flex h-7 items-center gap-3 border-t border-divider bg-background px-3 text-[11px] text-secondary"
      role="status"
    >
      <div className="flex min-w-0 items-center gap-2">{leftSlot}</div>
      <div
        className="flex min-w-0 flex-1 items-center justify-center gap-3 truncate"
        aria-live="polite"
        aria-atomic="true"
        data-testid="shell-selection-summary"
      >
        {summary?.aggregates && summary.aggregates.length > 0 ? (
          summary.aggregates.map((a) => (
            <span key={a.label} className="tabular-nums">
              <span className="text-tertiary">{a.label}: </span>
              <span className="text-foreground">{a.value}</span>
            </span>
          ))
        ) : summary?.text ? (
          <span className="truncate">{summary.text}</span>
        ) : null}
      </div>
      <div className="flex items-center gap-2">
        {rightSlot}
        <SaveStateMirror state={adapter.saveState} />
      </div>
    </footer>
  );
}

function SaveStateMirror({ state }: { readonly state: SaveState }): ReactNode {
  if (state === "unknown") return null;
  const label =
    state === "saved"
      ? "Saved"
      : state === "modified"
        ? "Unsaved changes"
        : state === "saving"
          ? "Saving…"
          : "Save error";
  const tone =
    state === "error"
      ? "text-[color:var(--error)]"
      : state === "modified"
        ? "text-[color:var(--warning)]"
        : "text-tertiary";
  return <span className={cn("tabular-nums", tone)}>{label}</span>;
}
