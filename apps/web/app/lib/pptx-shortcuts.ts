"use client";

import { useEffect } from "react";
import type { PptxAgent, Shape } from "@officeai/pptx";
import { isMixed, type TextFormatProvider } from "@officeai/text-formatting";

/**
 * PPTX keyboard shortcuts. First keyboard layer in this product, so
 * the design intentionally stays tiny: a single `keydown` listener on
 * the slide-canvas root that dispatches PPT-flavoured commands via
 * the existing agent / provider plumbing.
 *
 * Bindings (Cmd on macOS, Ctrl elsewhere):
 *
 *   Selected shape:
 *     Backspace / Delete              → delete shape (already wired to
 *                                       the toolbar — we just give it a
 *                                       keyboard route).
 *     Arrow keys                      → nudge by 1 px (≈ 9525 EMU).
 *     Shift + Arrow keys              → move by 10 px.
 *     Mod-B / Mod-I / Mod-U           → toggle bold/italic/underline
 *                                       on the active text selection
 *                                       (no-op when no text run is
 *                                       selected; falls through so the
 *                                       browser's own keymap isn't
 *                                       hijacked unnecessarily).
 *
 *   Slide-level:
 *     Mod-M                           → add new slide.
 *     Mod-Shift-D                     → duplicate current slide.
 *     PageUp / PageDown               → previous / next slide.
 *
 * Skipped entirely when the contenteditable text overlay owns focus —
 * typing inside a shape stays unaffected, except for Mod-B/I/U which
 * is the explicit override for that surface.
 */
export interface PptxShortcutDeps {
  /** The slide canvas surface; the listener is scoped to it via window-level capture. */
  readonly surfaceRef: { readonly current: HTMLElement | null };
  readonly agentRef: { readonly current: PptxAgent | null };
  readonly activeIndex: number;
  readonly slideCount: number;
  readonly selectedShape: Shape | null;
  readonly selectedShapeIds: ReadonlyArray<string>;
  readonly textFormatProvider: TextFormatProvider;
  readonly onAddSlide: () => void;
  readonly onDuplicateSlide: () => void;
  readonly onDeleteShape: () => void;
  readonly onChangeSlide: (index: number) => void;
  readonly onError: (err: unknown) => void;
}

/** ≈ 1 px in EMU at 96 DPI (1 inch = 914400 EMU = 96 px). */
const PX_TO_EMU = 9525;

export function usePptxShortcuts(deps: PptxShortcutDeps): void {
  // We deliberately resubscribe whenever any dep changes — there's
  // exactly one listener and the closure must read the freshest
  // selection / index, which would otherwise stale-close on the
  // initial mount.
  useEffect(() => {
    function onKeyDown(event: KeyboardEvent) {
      const surface = deps.surfaceRef.current;
      if (!surface) return;
      const target = event.target as Node | null;
      // Only handle keys that originated inside the slide surface so
      // we don't fight other editors on the same page.
      if (!target || !(surface === target || surface.contains(target))) return;

      const isMod = event.metaKey || event.ctrlKey;
      const shift = event.shiftKey;
      const alt = event.altKey;
      const editingText = isContentEditableTarget(target);

      // Mod+B/I/U work BOTH when editing text (apply to selected runs)
      // and when only shape-selected (no-op since provider has no
      // selection — `hasSelection()` returns false).
      if (isMod && !shift && !alt) {
        const mark = matchInlineMark(event);
        if (mark) {
          if (!deps.textFormatProvider.hasSelection()) return;
          event.preventDefault();
          try {
            deps.textFormatProvider.apply({ [mark]: !readActiveBool(deps.textFormatProvider, mark) });
          } catch (err) {
            deps.onError(err);
          }
          return;
        }
      }

      // Anything below targets the *shape* selection. Bail when the
      // user is typing inside a text frame — those keys belong to the
      // contenteditable overlay.
      if (editingText) return;

      // Slide-level navigation works regardless of shape selection.
      if (event.key === "PageDown" && !isMod && !shift && !alt) {
        if (deps.activeIndex < deps.slideCount - 1) {
          event.preventDefault();
          deps.onChangeSlide(deps.activeIndex + 1);
        }
        return;
      }
      if (event.key === "PageUp" && !isMod && !shift && !alt) {
        if (deps.activeIndex > 0) {
          event.preventDefault();
          deps.onChangeSlide(deps.activeIndex - 1);
        }
        return;
      }

      if (isMod && !shift && !alt && (event.key === "m" || event.key === "M")) {
        event.preventDefault();
        deps.onAddSlide();
        return;
      }
      if (isMod && shift && !alt && (event.key === "d" || event.key === "D")) {
        event.preventDefault();
        deps.onDuplicateSlide();
        return;
      }

      // Shape-level operations require a selection.
      if (deps.selectedShapeIds.length === 0) return;

      if (event.key === "Backspace" || event.key === "Delete") {
        if (isMod || shift || alt) return;
        event.preventDefault();
        deps.onDeleteShape();
        return;
      }

      const nudge = arrowDelta(event.key);
      if (nudge) {
        if (isMod || alt) return;
        event.preventDefault();
        const factor = shift ? 10 : 1;
        void moveSelection(deps, nudge[0] * factor, nudge[1] * factor);
        return;
      }
    }

    window.addEventListener("keydown", onKeyDown, true);
    return () => window.removeEventListener("keydown", onKeyDown, true);
  }, [deps]);
}

function isContentEditableTarget(target: Node | null): boolean {
  if (!target) return false;
  let el: HTMLElement | null =
    target.nodeType === Node.ELEMENT_NODE ? (target as HTMLElement) : (target.parentElement ?? null);
  while (el) {
    if (el.isContentEditable) return true;
    el = el.parentElement;
  }
  return false;
}

function matchInlineMark(event: KeyboardEvent): "bold" | "italic" | "underline" | null {
  const k = event.key;
  if (k === "b" || k === "B") return "bold";
  if (k === "i" || k === "I") return "italic";
  if (k === "u" || k === "U") return "underline";
  return null;
}

function arrowDelta(key: string): readonly [number, number] | null {
  switch (key) {
    case "ArrowUp":
      return [0, -1];
    case "ArrowDown":
      return [0, 1];
    case "ArrowLeft":
      return [-1, 0];
    case "ArrowRight":
      return [1, 0];
    default:
      return null;
  }
}

function readActiveBool(provider: TextFormatProvider, mark: "bold" | "italic" | "underline"): boolean {
  const active = provider.getActive();
  const v = active[mark];
  // Treat MIXED or any truthy value (including non-default underline
  // styles like "wavy") as "currently on" so a second press clears
  // the mark across the selection (mirrors Word's behaviour).
  if (isMixed(v)) return true;
  return Boolean(v);
}

async function moveSelection(deps: PptxShortcutDeps, dxPx: number, dyPx: number): Promise<void> {
  const a = deps.agentRef.current;
  const shape = deps.selectedShape;
  if (!a || !shape || shape.position === undefined) return;
  const x = shape.position.xEmu + dxPx * PX_TO_EMU;
  const y = shape.position.yEmu + dyPx * PX_TO_EMU;
  try {
    await a.applyCommand({
      type: "pptx:set-position",
      payload: { slideIndex: deps.activeIndex, shapeId: shape.id, x, y },
      source: "human",
    });
  } catch (err) {
    deps.onError(err);
  }
}
