"use client";

import {
  cloneElement,
  isValidElement,
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type ReactElement,
  type ReactNode,
} from "react";
import { createPortal } from "react-dom";
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
 * pickers. Closes on outside click or Escape.
 *
 * The panel is portalled into `document.body` and positioned with
 * `position: fixed` from the trigger's bounding rect. This is the
 * same pattern as `ToolbarMenu` in the web app and exists for the
 * same reason: the shared `ToolbarRow` puts its leading slot inside
 * `overflow-x-auto overflow-y-hidden` so a wide toolbar scrolls
 * horizontally instead of compressing button hit-targets. Vertically,
 * that same `overflow-y-hidden` clips any absolutely-positioned panel
 * extending below the 40 px row — which made the font color and
 * highlight palettes appear to render but the swatches below the
 * first row of pixels were unreachable (clicks fell through to the
 * canvas/grid below). Portalling out of the toolbar fixes both the
 * visual clipping and hit-testing.
 */
export function Popover({ trigger, children, panelClassName, align = "start" }: PopoverProps): ReactNode {
  const [open, setOpen] = useState(false);
  const triggerWrapperRef = useRef<HTMLDivElement | null>(null);
  const panelRef = useRef<HTMLDivElement | null>(null);
  const [pos, setPos] = useState<{ top: number; left: number } | null>(null);

  const reposition = useCallback(() => {
    const wrapper = triggerWrapperRef.current;
    const panel = panelRef.current;
    if (!wrapper || !panel) return;
    const rect = wrapper.getBoundingClientRect();
    const panelRect = panel.getBoundingClientRect();
    let left = align === "end" ? rect.right - panelRect.width : rect.left;
    // Clamp inside the viewport with an 8 px gutter on each side so
    // a wide panel near a screen edge still reads cleanly.
    const gutter = 8;
    const maxLeft = window.innerWidth - panelRect.width - gutter;
    if (left > maxLeft) left = Math.max(gutter, maxLeft);
    if (left < gutter) left = gutter;
    const top = rect.bottom + 4;
    setPos({ top, left });
  }, [align]);

  useLayoutEffect(() => {
    if (!open) {
      setPos(null);
      return;
    }
    reposition();
    const onScroll = () => reposition();
    const onResize = () => reposition();
    window.addEventListener("resize", onResize);
    // Capture-phase so we hear about scrolls in any ancestor scroller
    // (the toolbar's overflow-x strip is the relevant one).
    window.addEventListener("scroll", onScroll, true);
    return () => {
      window.removeEventListener("resize", onResize);
      window.removeEventListener("scroll", onScroll, true);
    };
  }, [open, reposition]);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => {
      const target = e.target as Node | null;
      if (!target) return;
      if (panelRef.current?.contains(target)) return;
      if (triggerWrapperRef.current?.contains(target)) return;
      setOpen(false);
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

  const panel =
    open && typeof document !== "undefined"
      ? createPortal(
          <div
            ref={panelRef}
            role="menu"
            className={cn("rounded-md border border-divider bg-surface p-2 shadow-md", panelClassName)}
            style={{
              position: "fixed",
              top: pos?.top ?? -9999,
              left: pos?.left ?? -9999,
              // Hide until the first measurement lands, otherwise the
              // panel briefly flashes at (0, 0) before reposition() runs.
              visibility: pos ? "visible" : "hidden",
              zIndex: 60,
            }}
            onClick={(e) => {
              // Close after any click inside the panel that isn't explicitly stopped.
              const target = e.target as HTMLElement;
              if (target.closest("[data-popover-close='false']")) return;
              setOpen(false);
            }}
          >
            {children}
          </div>,
          document.body
        )
      : null;

  return (
    <div ref={triggerWrapperRef} className="relative inline-flex">
      {enhancedTrigger}
      {panel}
    </div>
  );
}
