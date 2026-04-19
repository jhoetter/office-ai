"use client";

import {
  cloneElement,
  isValidElement,
  useEffect,
  useRef,
  useState,
  type ReactElement,
  type ReactNode,
} from "react";
import { cn } from "../lib/cn";

export interface PopoverProps {
  /**
   * The trigger element. Receives `aria-expanded` and an `onClick`
   * handler that toggles the popover.
   */
  trigger: ReactElement;
  /** Popover contents. Rendered inside an absolutely-positioned panel. */
  children: ReactNode;
  /** Optional className for the panel. */
  panelClassName?: string;
  /** Aligns the panel to the trigger. Default `start`. */
  align?: "start" | "end";
}

/**
 * Minimal accessible popover used by the shared text-formatting
 * pickers. Closes on outside click or Escape. No portal — the panel
 * is positioned absolute relative to the wrapper div.
 */
export function Popover({ trigger, children, panelClassName, align = "start" }: PopoverProps): ReactNode {
  const [open, setOpen] = useState(false);
  const wrapperRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => {
      if (!wrapperRef.current) return;
      if (!wrapperRef.current.contains(e.target as Node)) setOpen(false);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setOpen(false);
    };
    document.addEventListener("mousedown", onDoc);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDoc);
      document.removeEventListener("keydown", onKey);
    };
  }, [open]);

  if (
    !isValidElement<{
      "aria-expanded"?: boolean;
      "aria-haspopup"?: string;
      onClick?: (e: React.MouseEvent) => void;
    }>(trigger)
  ) {
    throw new Error("Popover requires a valid React element as trigger.");
  }

  const enhancedTrigger = cloneElement(trigger, {
    "aria-expanded": open,
    "aria-haspopup": "menu",
    onClick: (e: React.MouseEvent) => {
      trigger.props.onClick?.(e);
      if (!e.defaultPrevented) setOpen((v) => !v);
    },
  });

  return (
    <div ref={wrapperRef} className="relative inline-flex">
      {enhancedTrigger}
      {open && (
        <div
          role="menu"
          className={cn(
            "absolute top-full z-30 mt-1 rounded-md border border-divider bg-surface p-2 shadow-md",
            align === "end" ? "right-0" : "left-0",
            panelClassName
          )}
          onClick={(e) => {
            // Close after any click inside the panel that isn't explicitly stopped.
            const target = e.target as HTMLElement;
            if (target.closest("[data-popover-close='false']")) return;
            setOpen(false);
          }}
        >
          {children}
        </div>
      )}
    </div>
  );
}
