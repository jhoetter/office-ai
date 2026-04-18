"use client";

import { type ReactNode } from "react";
import {
  FONT_FAMILIES,
  MIXED,
  type MaybeMixed,
} from "@officeai/text-formatting";

export interface FontFamilyPickerProps {
  value: MaybeMixed<string>;
  onChange: (family: string) => void;
  disabled?: boolean;
  /** Override the curated list. */
  families?: ReadonlyArray<string>;
  className?: string;
}

/**
 * Native <select>-based font-family picker. The active value is
 * always present in the option list — even when it's not in the
 * curated FONT_FAMILIES list — so a doc with "Aptos" still shows
 * "Aptos" as selected.
 */
export function FontFamilyPicker({
  value,
  onChange,
  disabled,
  families = FONT_FAMILIES,
  className,
}: FontFamilyPickerProps): ReactNode {
  const concreteValue = typeof value === "string" ? value : "";
  const augmented =
    concreteValue && !families.includes(concreteValue)
      ? [concreteValue, ...families]
      : families;
  const placeholder = value === MIXED ? "—" : "Font";
  return (
    <label className="inline-flex items-center gap-1 text-xs text-secondary">
      <span className="sr-only">Font family</span>
      <select
        title="Font family"
        aria-label="Font family"
        value={concreteValue}
        disabled={disabled}
        onChange={(e) => {
          if (!e.target.value) return;
          onChange(e.target.value);
        }}
        className={
          className ??
          "h-7 w-32 rounded-md border border-divider bg-surface px-2 text-xs text-foreground hover:bg-hover focus:outline-none"
        }
      >
        <option value="" disabled>
          {placeholder}
        </option>
        {augmented.map((f) => (
          <option key={f} value={f}>
            {f}
          </option>
        ))}
      </select>
    </label>
  );
}
