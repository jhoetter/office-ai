"use client";

import * as React from "react";
import { createPortal } from "react-dom";

/**
 * Portal-rendered popover for toolbar dropdown menus.
 *
 * Why a portal: the shared `ToolbarRow` puts its leading slot inside
 * `overflow-x-auto overflow-y-hidden` so a wide toolbar scrolls
 * horizontally instead of compressing button hit-targets. Vertically,
 * that same `overflow` clips any absolutely-positioned dropdown that
 * tries to extend below the 40 px row, which makes menu items appear
 * to render but clicks then fall through to the slide canvas (the
 * actual DOM element under the cursor is the slide's SVG, not the
 * menu item). Rendering the panel into `document.body` with
 * `position: fixed` escapes the clip entirely and keeps the menu
 * visible and clickable.
 *
 * The component is intentionally minimal:
 *
 *  - Position is computed from the trigger's `getBoundingClientRect`
 *    on open + on resize/scroll, so the panel tracks the trigger
 *    even when the toolbar's horizontal scroll moves it.
 *  - Click-outside dismissal honours both the trigger and the panel
 *    so clicking inside the panel doesn't close it.
 *  - Escape closes the panel — matching every other menu in the app.
 */
export interface ToolbarMenuProps {
  /** Open state — owned by the parent so callers can close on
   * pick. */
  readonly open: boolean;
  readonly onClose: () => void;
  /** Ref to the trigger button. Used to anchor the panel and to
   * exempt the trigger from click-outside dismissal. */
  readonly triggerRef: React.RefObject<HTMLElement | null>;
  /**
   * Horizontal alignment relative to the trigger. `left` lines the
   * panel's left edge with the trigger's left edge (the default,
   * matches every existing toolbar dropdown), `right` lines the
   * right edges (use for trailing menus that would otherwise
   * overflow off-screen).
   */
  readonly align?: "left" | "right";
  /** Pixels between the trigger and the panel. Defaults to 4. */
  readonly offset?: number;
  readonly className?: string;
  readonly children: React.ReactNode;
  /** Forwarded straight to the panel `<div>` for tests / a11y. */
  readonly role?: string;
  readonly testId?: string;
}

export function ToolbarMenu({
  open,
  onClose,
  triggerRef,
  align = "left",
  offset = 4,
  className,
  children,
  role,
  testId,
}: ToolbarMenuProps): React.ReactElement | null {
  const panelRef = React.useRef<HTMLDivElement | null>(null);
  const [pos, setPos] = React.useState<{ top: number; left: number } | null>(null);

  // Recompute panel coordinates from the trigger rect. Called on
  // open, on window resize, and on scroll (capture-phase, so we
  // catch scrolls inside the toolbar's own overflow-x-auto strip).
  const reposition = React.useCallback(() => {
    const trigger = triggerRef.current;
    const panel = panelRef.current;
    if (!trigger || !panel) return;
    const rect = trigger.getBoundingClientRect();
    const panelRect = panel.getBoundingClientRect();
    let left =
      align === "right" ? rect.right - panelRect.width : rect.left;
    // Keep the panel inside the viewport — clamp with an 8 px gutter
    // on each side so a long menu near a screen edge still reads
    // cleanly.
    const gutter = 8;
    const maxLeft = window.innerWidth - panelRect.width - gutter;
    if (left > maxLeft) left = Math.max(gutter, maxLeft);
    if (left < gutter) left = gutter;
    const top = rect.bottom + offset;
    setPos({ top, left });
  }, [align, offset, triggerRef]);

  React.useLayoutEffect(() => {
    if (!open) {
      setPos(null);
      return;
    }
    reposition();
    const onScroll = () => reposition();
    const onResize = () => reposition();
    window.addEventListener("resize", onResize);
    // Capture-phase so we hear about scrolls in any ancestor
    // scroller (the toolbar's overflow-x strip is the relevant
    // one).
    window.addEventListener("scroll", onScroll, true);
    return () => {
      window.removeEventListener("resize", onResize);
      window.removeEventListener("scroll", onScroll, true);
    };
  }, [open, reposition]);

  React.useEffect(() => {
    if (!open) return;
    function onDoc(e: MouseEvent) {
      const target = e.target as Node | null;
      if (!target) return;
      if (panelRef.current?.contains(target)) return;
      if (triggerRef.current?.contains(target)) return;
      onClose();
    }
    function onKey(e: KeyboardEvent) {
      if (e.key === "Escape") onClose();
    }
    document.addEventListener("mousedown", onDoc);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDoc);
      document.removeEventListener("keydown", onKey);
    };
  }, [open, onClose, triggerRef]);

  if (!open) return null;
  if (typeof document === "undefined") return null;

  return createPortal(
    <div
      ref={panelRef}
      role={role}
      data-testid={testId}
      className={className}
      style={{
        position: "fixed",
        top: pos?.top ?? -9999,
        left: pos?.left ?? -9999,
        // Hide until the first measurement lands, otherwise the
        // panel briefly flashes at (0, 0) before reposition() runs.
        visibility: pos ? "visible" : "hidden",
        zIndex: 60,
      }}
    >
      {children}
    </div>,
    document.body
  );
}
