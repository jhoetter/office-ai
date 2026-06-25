"use client";

import { type ReactNode } from "react";
import { ChevronDown } from "../sonaloop-icons";
import {
  COLOR_PALETTE,
  MIXED,
  renderColor,
  type ColorSwatch,
  type MaybeMixed,
} from "@officeai/text-formatting";
import { Popover } from "./popover";

export interface ColorPickerProps {
  label: string;
  /** Icon shown on the trigger button. */
  icon: ReactNode;
  /** Lowercase RRGGBB or `MIXED` / `undefined`. */
  value: MaybeMixed<string>;
  /** Called with the new color (lowercase RRGGBB, no '#'). */
  onChange: (rrggbb: string) => void;
  /** Optional "clear" callback — when omitted, no clear affordance is rendered. */
  onClear?: () => void;
  /** Override the curated palette. */
  palette?: ReadonlyArray<ColorSwatch>;
  disabled?: boolean;
}

/**
 * Color picker dropdown — palette grid + optional "clear" affordance.
 * Reuses the shared Popover; the active color is shown as a thin
 * stripe under the icon (matching Word's UX), `MIXED` renders as a
 * neutral stripe with a `—` overlay tooltip.
 */
export function ColorPicker({
  label,
  icon,
  value,
  onChange,
  onClear,
  palette = COLOR_PALETTE,
  disabled,
}: ColorPickerProps): ReactNode {
  const stripeColor =
    value === MIXED ? "var(--divider)" : typeof value === "string" ? renderColor(value) : undefined;

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
          Clear
        </button>
      )}
    </Popover>
  );
}
