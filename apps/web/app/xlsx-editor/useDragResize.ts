"use client";

import { useEffect, useRef, useState } from "react";
import { bodyPxToAnchor, type AnchorFromPx } from "./ImageOverlay";
import type { AxisLookup } from "./gridDimensions";
import type { ResizeHandleGrabInfo } from "@officeai/ui";

/**
 * Shared drag + 8-handle resize state machine for any "free-floating
 * thing anchored to the grid" — currently used by both
 * {@link ImageOverlay} and {@link ChartOverlay}. Lives outside the
 * overlay components so the two callers can't drift apart in subtle
 * ways (e.g. snapping rules, min-size, "did the move actually
 * change anything" comparison).
 *
 * The hook returns:
 *   - the *display* rectangle (committed values when idle, transient
 *     values while dragging)
 *   - `startMove` / `startResize` mouse-down handlers to wire onto
 *     the body and each handle
 *
 * Commit happens on mouse-up via the supplied callbacks. The
 * callbacks are only invoked when the values actually changed, so
 * accidental click-without-drag doesn't push a no-op onto the
 * command bus / undo stack.
 */
type AxisLike = ReadonlyArray<number> | AxisLookup;

export type ResizeHandle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

type DragMode =
  | { kind: "move"; startMouseX: number; startMouseY: number; startBodyX: number; startBodyY: number }
  | {
      kind: "resize";
      handle: ResizeHandle;
      /**
       * Aspect ratio captured at gesture start so Shift-drag on a
       * corner handle can lock to it without re-reading the rect
       * each move. Mirrors the same field on `ImageOverlay`'s drag
       * mode — kept private to the hook so callers don't need to
       * thread it through.
       */
      aspect: number;
      startMouseX: number;
      startMouseY: number;
      startBodyX: number;
      startBodyY: number;
      startWidth: number;
      startHeight: number;
    };

export interface DragResizeAnchor {
  readonly fromRow: number;
  readonly fromCol: number;
  readonly fromOffsetXPx: number;
  readonly fromOffsetYPx: number;
  readonly widthPx: number;
  readonly heightPx: number;
}

export interface UseDragResizeArgs {
  readonly anchor: DragResizeAnchor;
  readonly baseBodyX: number;
  readonly baseBodyY: number;
  readonly colXs: AxisLike;
  readonly rowYs: AxisLike;
  readonly minWidth?: number;
  readonly minHeight?: number;
  readonly onMoveCommit: (anchor: AnchorFromPx) => void;
  readonly onResizeCommit: (size: { widthPx: number; heightPx: number }) => void;
}

export interface UseDragResizeResult {
  readonly displayBodyX: number;
  readonly displayBodyY: number;
  readonly displayWidth: number;
  readonly displayHeight: number;
  readonly dragKind: DragMode["kind"] | null;
  startMove(e: React.MouseEvent): void;
  /**
   * Begin a resize gesture. Accepts a {@link ResizeHandleGrabInfo}
   * payload as emitted by the shared `ResizeHandles` primitive so
   * the hook stays decoupled from React's synthetic event type.
   */
  startResize(info: ResizeHandleGrabInfo): void;
}

export function useDragResize(args: UseDragResizeArgs): UseDragResizeResult {
  const { anchor, baseBodyX, baseBodyY, colXs, rowYs, onMoveCommit, onResizeCommit } = args;
  const minWidth = args.minWidth ?? 8;
  const minHeight = args.minHeight ?? 8;

  const [dragMode, setDragMode] = useState<DragMode | null>(null);
  const [transient, setTransient] = useState<{
    bodyX: number;
    bodyY: number;
    widthPx: number;
    heightPx: number;
  } | null>(null);
  const transientRef = useRef(transient);
  useEffect(() => {
    transientRef.current = transient;
  }, [transient]);

  useEffect(() => {
    if (!dragMode) return;
    const onMove = (e: MouseEvent) => {
      if (dragMode.kind === "move") {
        const dx = e.clientX - dragMode.startMouseX;
        const dy = e.clientY - dragMode.startMouseY;
        setTransient({
          bodyX: Math.max(0, dragMode.startBodyX + dx),
          bodyY: Math.max(0, dragMode.startBodyY + dy),
          widthPx: anchor.widthPx,
          heightPx: anchor.heightPx,
        });
        return;
      }
      const dx = e.clientX - dragMode.startMouseX;
      const dy = e.clientY - dragMode.startMouseY;
      // Read Shift live every move (matches `ImageOverlay`) — letting
      // the user toggle aspect lock without re-grabbing.
      const proportional = e.shiftKey && isCorner(dragMode.handle);
      setTransient(applyResize(dragMode, dx, dy, minWidth, minHeight, proportional));
    };
    const onUp = () => {
      const final = transientRef.current;
      const mode = dragMode;
      setDragMode(null);
      setTransient(null);
      if (!final) return;
      if (mode.kind === "move") {
        const next = bodyPxToAnchor(final.bodyX, final.bodyY, colXs, rowYs);
        if (anchorChanged(next, anchor)) onMoveCommit(next);
      } else {
        const widthPx = Math.max(minWidth, Math.round(final.widthPx));
        const heightPx = Math.max(minHeight, Math.round(final.heightPx));
        if (widthPx !== anchor.widthPx || heightPx !== anchor.heightPx) {
          onResizeCommit({ widthPx, heightPx });
        }
        if (mode.handle.includes("w") || mode.handle.includes("n")) {
          const next = bodyPxToAnchor(final.bodyX, final.bodyY, colXs, rowYs);
          if (anchorChanged(next, anchor)) onMoveCommit(next);
        }
      }
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
  }, [dragMode, colXs, rowYs, anchor, onMoveCommit, onResizeCommit, minWidth, minHeight]);

  const displayBodyX = transient?.bodyX ?? baseBodyX;
  const displayBodyY = transient?.bodyY ?? baseBodyY;
  const displayWidth = transient?.widthPx ?? anchor.widthPx;
  const displayHeight = transient?.heightPx ?? anchor.heightPx;

  return {
    displayBodyX,
    displayBodyY,
    displayWidth,
    displayHeight,
    dragKind: dragMode?.kind ?? null,
    startMove(e) {
      e.preventDefault();
      e.stopPropagation();
      setTransient({
        bodyX: baseBodyX,
        bodyY: baseBodyY,
        widthPx: anchor.widthPx,
        heightPx: anchor.heightPx,
      });
      setDragMode({
        kind: "move",
        startMouseX: e.clientX,
        startMouseY: e.clientY,
        startBodyX: baseBodyX,
        startBodyY: baseBodyY,
      });
    },
    startResize(info) {
      // The primitive already called preventDefault / stopPropagation
      // on the underlying mousedown, so callers don't have to.
      setTransient({
        bodyX: baseBodyX,
        bodyY: baseBodyY,
        widthPx: anchor.widthPx,
        heightPx: anchor.heightPx,
      });
      setDragMode({
        kind: "resize",
        handle: info.side,
        aspect: anchor.heightPx > 0 ? anchor.widthPx / anchor.heightPx : 1,
        startMouseX: info.clientX,
        startMouseY: info.clientY,
        startBodyX: baseBodyX,
        startBodyY: baseBodyY,
        startWidth: anchor.widthPx,
        startHeight: anchor.heightPx,
      });
    },
  };
}

function isCorner(h: ResizeHandle): boolean {
  return h === "nw" || h === "ne" || h === "se" || h === "sw";
}

function applyResize(
  drag: Extract<DragMode, { kind: "resize" }>,
  dx: number,
  dy: number,
  minW: number,
  minH: number,
  proportional: boolean
): { bodyX: number; bodyY: number; widthPx: number; heightPx: number } {
  let { startBodyX: x, startBodyY: y, startWidth: w, startHeight: h } = drag;
  if (drag.handle.includes("e")) w = Math.max(minW, drag.startWidth + dx);
  if (drag.handle.includes("s")) h = Math.max(minH, drag.startHeight + dy);
  if (drag.handle.includes("w")) {
    const newW = Math.max(minW, drag.startWidth - dx);
    x = drag.startBodyX + (drag.startWidth - newW);
    w = newW;
  }
  if (drag.handle.includes("n")) {
    const newH = Math.max(minH, drag.startHeight - dy);
    y = drag.startBodyY + (drag.startHeight - newH);
    h = newH;
  }
  if (proportional && isCorner(drag.handle) && drag.aspect > 0) {
    const rW = w / drag.startWidth;
    const rH = h / drag.startHeight;
    const widthDriven = Math.abs(rW - 1) >= Math.abs(rH - 1);
    if (widthDriven) {
      h = Math.max(minH, w / drag.aspect);
    } else {
      w = Math.max(minW, h * drag.aspect);
    }
    if (drag.handle.includes("w")) x = drag.startBodyX + (drag.startWidth - w);
    if (drag.handle.includes("n")) y = drag.startBodyY + (drag.startHeight - h);
  }
  return { bodyX: Math.max(0, x), bodyY: Math.max(0, y), widthPx: w, heightPx: h };
}

function anchorChanged(next: AnchorFromPx, prev: DragResizeAnchor): boolean {
  return (
    next.fromRow !== prev.fromRow ||
    next.fromCol !== prev.fromCol ||
    Math.round(next.fromOffsetXPx) !== Math.round(prev.fromOffsetXPx) ||
    Math.round(next.fromOffsetYPx) !== Math.round(prev.fromOffsetYPx)
  );
}
