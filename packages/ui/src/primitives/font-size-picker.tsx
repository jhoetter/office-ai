"use client";

import { type ReactNode } from "react";
import {
  FONT_SIZES_PT,
  MIXED,
  type MaybeMixed,
} from "@officeai/text-formatting";

export interface FontSizePickerProps {
  /** Active size in points. */
  value: MaybeMixed<number>;
  /** Called with the new size in points. */
  onChange: (pt: number) => void;
  disabled?: boolean;
  /** Override the curated preset list (in points). */
  sizes?: ReadonlyArray<number>;
  className?: string;
}

/**
 * Font-size picker. Speaks points everywhere — adapters convert into
 * the format-specific unit (half-points / hundredths) on dispatch.
 */
export function FontSizePicker({
  value,
  onChange,
  disabled,
  sizes = FONT_SIZES_PT,
  className,
}: FontSizePickerProps): ReactNode {
  const display =
    value === MIXED
      ? "—"
      : typeof value === "number"
      ? formatPt(value)
      : "Size";
  const concreteValue =
    typeof value === "number" && sizes.includes(value) ? String(value) : "";
  return (
    <label className="inline-flex items-center gap-1 text-xs text-secondary">
      <span className="sr-only">Font size</span>
      <select
        title="Font size"
        aria-label="Font size"
        value={concreteValue}
        disabled={disabled}
        onChange={(e) => {
          const n = Number(e.target.value);
          if (!Number.isFinite(n) || n <= 0) return;
          onChange(n);
        }}
        className={
          className ??
          "h-7 w-16 rounded-md border border-divider bg-surface px-2 text-xs text-foreground hover:bg-hover focus:outline-none"
        }
      >
        <option value="" disabled>
          {display}
        </option>
        {sizes.map((pt) => (
          <option key={pt} value={pt}>
            {formatPt(pt)}
          </option>
        ))}
      </select>
    </label>
  );
}

function formatPt(pt: number): string {
  return Number.isInteger(pt) ? String(pt) : pt.toFixed(1);
}
