"use client";

import { useEffect, type RefObject } from "react";

const FOCUSABLE = [
  "a[href]",
  "button:not([disabled])",
  "input:not([disabled]):not([type='hidden'])",
  "select:not([disabled])",
  "textarea:not([disabled])",
  "[tabindex]:not([tabindex='-1'])",
].join(",");

/**
 * Trap focus inside a container while it's mounted, restore focus to
 * the previously-focused element on unmount, and forward Escape to an
 * optional close handler.
 *
 * Designed for modal dialogs across DOCX/XLSX/PPTX so they all behave
 * the same way: Tab cycles within the dialog, Shift+Tab cycles in
 * reverse, Escape closes, and the trigger button regains focus when
 * the dialog goes away. Pass `enabled=false` to disable without
 * unmounting (useful for non-modal popovers in modal-like containers).
 */
export function useFocusTrap(
  containerRef: RefObject<HTMLElement | null>,
  options: {
    readonly enabled?: boolean;
    readonly initialFocusRef?: RefObject<HTMLElement | null>;
    readonly onEscape?: () => void;
  } = {}
): void {
  const { enabled = true, initialFocusRef, onEscape } = options;

  useEffect(() => {
    if (!enabled) return;
    const container = containerRef.current;
    if (!container) return;

    const previouslyFocused = document.activeElement as HTMLElement | null;
    const initial = initialFocusRef?.current ?? container.querySelector<HTMLElement>(FOCUSABLE) ?? container;
    initial.focus();

    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === "Escape" && onEscape) {
        e.stopPropagation();
        e.preventDefault();
        onEscape();
        return;
      }
      if (e.key !== "Tab") return;
      const focusables = Array.from(container.querySelectorAll<HTMLElement>(FOCUSABLE)).filter(
        (el) => !el.hasAttribute("data-focus-skip")
      );
      if (focusables.length === 0) {
        e.preventDefault();
        container.focus();
        return;
      }
      const first = focusables[0];
      const last = focusables[focusables.length - 1];
      const active = document.activeElement as HTMLElement | null;
      if (e.shiftKey) {
        if (active === first || !container.contains(active)) {
          e.preventDefault();
          last.focus();
        }
      } else {
        if (active === last || !container.contains(active)) {
          e.preventDefault();
          first.focus();
        }
      }
    };

    container.addEventListener("keydown", onKeyDown);
    return () => {
      container.removeEventListener("keydown", onKeyDown);
      previouslyFocused?.focus?.();
    };
  }, [containerRef, enabled, initialFocusRef, onEscape]);
}
