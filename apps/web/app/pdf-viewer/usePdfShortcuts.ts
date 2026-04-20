"use client";

import { useEffect } from "react";

/**
 * Callback bag wired by the {@link PdfEditor}. Each callback maps to
 * one row in `/spec/pdf/keyboard-shortcuts.md`. All are optional so
 * the hook can also be used in tests / storybook with a partial set.
 */
export interface PdfShortcutHandlers {
  readonly nextPage?: () => void;
  readonly prevPage?: () => void;
  readonly firstPage?: () => void;
  readonly lastPage?: () => void;
  readonly zoomIn?: () => void;
  readonly zoomOut?: () => void;
  readonly fitWidth?: () => void;
  readonly fitPage?: () => void;
  readonly actualSize?: () => void;
  readonly openSearch?: () => void;
  readonly rotateClockwise?: () => void;
  readonly rotateCounterClockwise?: () => void;
}

/**
 * True when the keydown originated inside an `<input>` / `<textarea>`
 * / `contenteditable` host. We bail in that case so the user can
 * type "j" inside the find bar without the canvas eating it.
 */
function isFormField(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  const tag = target.tagName;
  if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return true;
  if (target.isContentEditable) return true;
  return false;
}

/**
 * Window-level keyboard handler for the PDF viewer.
 *
 * Mirrors the per-product `usePptxShortcuts` shape — bind once, read
 * the latest callbacks from a ref-stable closure, never tear down
 * mid-session. We deliberately re-bind when `handlers` changes so
 * each callback always sees the fresh `agent` / `currentPage` in
 * scope; the bindings are cheap (one window listener) and this keeps
 * the hook usable from a parent that doesn't want to wrap every
 * callback in `useCallback`.
 *
 * Shortcuts (subset of the spec — the rest is fired from the toolbar
 * directly):
 *
 *   PageDown / Space / j / ArrowRight / n   → next page
 *   PageUp   / Shift+Space / k / ArrowLeft  → previous page
 *   Home / g                                  → first page
 *   End  / G (Shift+g)                        → last page
 *   + / =  (with or without modifier)          → zoom in
 *   - / _                                      → zoom out
 *   0                                          → actual size (100 %)
 *   1                                          → fit width
 *   2                                          → fit page
 *   ] (Cmd/Ctrl) → rotate clockwise           , [ (Cmd/Ctrl) → counter
 *   /  or Cmd/Ctrl+F (handled by EditorShell)  → open find
 */
export function usePdfShortcuts(handlers: PdfShortcutHandlers): void {
  useEffect(() => {
    function onKey(e: KeyboardEvent): void {
      if (isFormField(e.target)) return;
      const mod = e.metaKey || e.ctrlKey;
      const key = e.key;

      // Modifier-bound chords first — these need to win over the
      // single-letter cases so Cmd+= zooms instead of being eaten by
      // a hypothetical "= → zoom in" branch.
      if (mod) {
        if (key === "]" || key === "}") {
          e.preventDefault();
          handlers.rotateClockwise?.();
          return;
        }
        if (key === "[" || key === "{") {
          e.preventDefault();
          handlers.rotateCounterClockwise?.();
          return;
        }
        if (key === "+" || key === "=") {
          e.preventDefault();
          handlers.zoomIn?.();
          return;
        }
        if (key === "-" || key === "_") {
          e.preventDefault();
          handlers.zoomOut?.();
          return;
        }
        if (key === "0") {
          e.preventDefault();
          handlers.actualSize?.();
          return;
        }
        // Cmd/Ctrl+F is owned by EditorShell; we don't intercept it.
        return;
      }

      switch (key) {
        case "PageDown":
        case "j":
        case "n":
        case "ArrowRight":
          e.preventDefault();
          handlers.nextPage?.();
          return;
        case " ":
          e.preventDefault();
          if (e.shiftKey) handlers.prevPage?.();
          else handlers.nextPage?.();
          return;
        case "PageUp":
        case "k":
        case "p":
        case "ArrowLeft":
          e.preventDefault();
          handlers.prevPage?.();
          return;
        case "Home":
        case "g":
          e.preventDefault();
          handlers.firstPage?.();
          return;
        case "End":
        case "G":
          e.preventDefault();
          handlers.lastPage?.();
          return;
        case "+":
        case "=":
          e.preventDefault();
          handlers.zoomIn?.();
          return;
        case "-":
        case "_":
          e.preventDefault();
          handlers.zoomOut?.();
          return;
        case "0":
          e.preventDefault();
          handlers.actualSize?.();
          return;
        case "1":
          e.preventDefault();
          handlers.fitWidth?.();
          return;
        case "2":
          e.preventDefault();
          handlers.fitPage?.();
          return;
        case "/":
          e.preventDefault();
          handlers.openSearch?.();
          return;
        case "[":
          e.preventDefault();
          handlers.rotateCounterClockwise?.();
          return;
        case "]":
          e.preventDefault();
          handlers.rotateClockwise?.();
          return;
        default:
          return;
      }
    }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [handlers]);
}
