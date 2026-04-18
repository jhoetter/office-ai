"use client";

import { type ReactNode } from "react";
import { ChevronDown } from "lucide-react";
import {
  HIGHLIGHT_PALETTE,
  MIXED,
  renderColor,
  type HighlightSwatch,
  type MaybeMixed,
} from "@officeai/text-formatting";
import { Popover } from "./popover";

export interface HighlightPickerProps {
  label: string;
  icon: ReactNode;
  /** Lowercase RRGGBB or `MIXED` / `undefined`. */
  value: MaybeMixed<string>;
  /** Called with the new highlight (lowercase RRGGBB, no '#'). */
  onChange: (rrggbb: string) => void;
  /** Clears the highlight — adapters typically dispatch `{ highlight: "" }`. */
  onClear?: () => void;
  /** Override the curated palette. */
  palette?: ReadonlyArray<HighlightSwatch>;
  disabled?: boolean;
}

/**
 * Highlight picker. Identical UX to ColorPicker but uses the
 * highlight-specific palette (which carries OOXML w:highlight enum
 * names for DOCX adapters to quantise onto).
 */
export function HighlightPicker({
  label,
  icon,
  value,
  onChange,
  onClear,
  palette = HIGHLIGHT_PALETTE,
  disabled,
}: HighlightPickerProps): ReactNode {
  const stripeColor =
    value === MIXED
      ? "var(--divider)"
      : typeof value === "string"
      ? renderColor(value)
      : undefined;

  return (
    <Popover
      panelClassName="w-44"
      trigger={
        <button
          type="button"
          title={label}
          aria-label={label}
          disabled={disabled}
          onMouseDown={(e) => e.preventDefault()}
          className="flex flex-col items-center gap-0 rounded-md p-1.5 text-secondary hover:bg-hover hover:text-foreground disabled:opacity-50"
        >
          <span className="flex items-center gap-0.5">
            {icon}
            <ChevronDown size={10} />
          </span>
          <span
            aria-hidden
            className="mt-0.5 block h-0.5 w-4 rounded-sm"
            style={{ background: stripeColor ?? "transparent" }}
          />
        </button>
      }
    >
      <div className="grid grid-cols-4 gap-1">
        {palette.map((swatch) => (
          <button
            key={swatch.hex}
            type="button"
            role="menuitem"
            title={swatch.name}
            aria-label={`${label}: ${swatch.name}`}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => onChange(swatch.hex)}
            className="flex h-7 w-7 items-center justify-center rounded border border-divider hover:scale-110"
            style={{ background: `#${swatch.hex}` }}
          />
        ))}
      </div>
      {onClear && (
        <button
          type="button"
          onMouseDown={(e) => e.preventDefault()}
          onClick={onClear}
          className="mt-2 w-full rounded-md border border-divider px-2 py-1 text-xs text-secondary hover:bg-hover hover:text-foreground"
        >
          No highlight
        </button>
      )}
    </Popover>
  );
}
