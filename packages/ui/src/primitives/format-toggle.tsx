"use client";

import { forwardRef, type ButtonHTMLAttributes, type ReactNode } from "react";
import { MIXED, isOnTruthy, type MaybeMixed } from "@officeai/text-formatting";
import { cn } from "../lib/cn";

export interface FormatToggleProps extends Omit<ButtonHTMLAttributes<HTMLButtonElement>, "value"> {
  label: string;
  /**
   * Active state for the toggle. `MIXED` is rendered as a slightly
   * dimmed pressed style so users see "some-but-not-all".
   */
  value: MaybeMixed<boolean>;
  children: ReactNode;
}

/**
 * Single B/I/U/S-style toggle. Click semantics are owned by the
 * caller — this component just renders a button reflecting the
 * `value` and forwards the click. Provider adapters typically pass
 * `() => apply({ bold: !isOnTruthy(value) })` so each click flips
 * the value relative to the current selection state, matching
 * Word/Excel/PowerPoint behaviour.
 */
export const FormatToggle = forwardRef<HTMLButtonElement, FormatToggleProps>(
  ({ label, value, className, children, onMouseDown, ...rest }, ref) => {
    const on = isOnTruthy(value);
    const mixed = value === MIXED;
    return (
      <button
        ref={ref}
        type="button"
        title={label}
        aria-label={label}
        aria-pressed={on}
        // Keep the editing surface (PM editor / contenteditable
        // overlay) focused so the live selection survives the click.
        // Without this, clicking Bold while editing a PPTX text shape
        // would blur the contenteditable, collapse the native
        // selection, and the dispatch would see "no selection".
        // Callers can still pass their own onMouseDown — we run it
        // first; if they preventDefault, we don't override.
        onMouseDown={(event) => {
          onMouseDown?.(event);
          if (!event.defaultPrevented) event.preventDefault();
        }}
        className={cn(
          "rounded-md p-1.5 text-secondary hover:bg-hover hover:text-foreground",
          on && "bg-accent-light text-foreground",
          mixed && "bg-accent-light/40 text-foreground",
          className
        )}
        {...rest}
      >
        {children}
      </button>
    );
  }
);
FormatToggle.displayName = "FormatToggle";
