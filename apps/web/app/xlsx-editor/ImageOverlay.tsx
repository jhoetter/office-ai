"use client";

import { useEffect, useRef, useState, type CSSProperties, type ReactNode } from "react";
import type { SheetImage } from "@officeai/xlsx";
import type { AxisLookup } from "./gridDimensions";

/**
 * Axis source consumed by the overlay. Either a real
 * `ReadonlyArray<number>` (legacy callers) or the lazy
 * {@link AxisLookup} proxy returned by {@link colXsView} /
 * {@link rowYsView}. Both expose `arr[i]` indexing and `arr.length`,
 * which is all the geometry helpers in this module need.
 */
type AxisLike = ReadonlyArray<number> | AxisLookup;

/**
 * Result of mapping a body-coordinate pixel back onto the cell grid.
 */
export interface AnchorFromPx {
  readonly fromRow: number;
  readonly fromCol: number;
  readonly fromOffsetXPx: number;
  readonly fromOffsetYPx: number;
}

/**
 * Convert an image's `(fromRow, fromCol, fromOffsetX, fromOffsetY)`
 * anchor into absolute body-coordinate pixels (i.e. relative to the
 * top-left of the cell area, EXCLUDING the row/column header bands).
 */
export function anchorToBodyPx(
  image: SheetImage,
  colXs: AxisLike,
  rowYs: AxisLike
): { x: number; y: number } {
  const c = clamp(image.anchor.fromCol, 0, colXs.length - 2);
  const r = clamp(image.anchor.fromRow, 0, rowYs.length - 2);
  return {
    x: (colXs[c] ?? 0) + image.anchor.fromOffsetXPx,
    y: (rowYs[r] ?? 0) + image.anchor.fromOffsetYPx,
  };
}

/**
 * Inverse of {@link anchorToBodyPx}: snap a body-coordinate pixel back
 * onto a `(fromRow, fromCol, offsetXPx, offsetYPx)` anchor by finding
 * the cell whose left/top edge sits at-or-before the pixel.
 */
export function bodyPxToAnchor(xPx: number, yPx: number, colXs: AxisLike, rowYs: AxisLike): AnchorFromPx {
  const x = Math.max(0, xPx);
  const y = Math.max(0, yPx);
  const fromCol = floorIndex(colXs, x);
  const fromRow = floorIndex(rowYs, y);
  return {
    fromRow,
    fromCol,
    fromOffsetXPx: x - (colXs[fromCol] ?? 0),
    fromOffsetYPx: y - (rowYs[fromRow] ?? 0),
  };
}

function clamp(v: number, lo: number, hi: number): number {
  return v < lo ? lo : v > hi ? hi : v;
}

/**
 * Largest index `i` such that `arr[i] <= target`. `arr` is the
 * monotonically-increasing prefix-sum (`colXs` / `rowYs`).
 */
function floorIndex(arr: AxisLike, target: number): number {
  let lo = 0;
  let hi = arr.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if ((arr[mid] ?? 0) <= target) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

const HANDLE_SIZE = 9;

type DragMode =
  | { kind: "move"; startMouseX: number; startMouseY: number; startBodyX: number; startBodyY: number }
  | {
      kind: "resize";
      handle: ResizeHandle;
      startMouseX: number;
      startMouseY: number;
      startBodyX: number;
      startBodyY: number;
      startWidth: number;
      startHeight: number;
    };

export type ResizeHandle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

const RESIZE_CURSORS: Record<ResizeHandle, string> = {
  nw: "nwse-resize",
  n: "ns-resize",
  ne: "nesw-resize",
  e: "ew-resize",
  se: "nwse-resize",
  s: "ns-resize",
  sw: "nesw-resize",
  w: "ew-resize",
};

export interface ImageOverlayProps {
  readonly image: SheetImage;
  readonly headerOffset: { x: number; y: number };
  readonly colXs: AxisLike;
  readonly rowYs: AxisLike;
  readonly src: string | undefined;
  readonly selected: boolean;
  readonly onSelect: () => void;
  readonly onMoveCommit: (anchor: AnchorFromPx) => void;
  readonly onResizeCommit: (size: { widthPx: number; heightPx: number }) => void;
  /**
   * Hit-test target the click handler should consult to mark the
   * outer DOM element with `data-image-id`. The XlsxEditor uses this
   * marker to suppress cell-mousedown handling when clicking through
   * to an image.
   */
  readonly imageId: string;
}

/**
 * Free-floating image that sits above the grid cells. Owns its own
 * pointer-driven drag and resize transient state; commits to the
 * parent (XlsxEditor → command bus) only on mouse-up so undo/redo
 * stays single-step per gesture.
 */
export function ImageOverlay(props: ImageOverlayProps): ReactNode {
  const {
    image,
    headerOffset,
    colXs,
    rowYs,
    src,
    selected,
    onSelect,
    onMoveCommit,
    onResizeCommit,
    imageId,
  } = props;

  const baseBodyPx = anchorToBodyPx(image, colXs, rowYs);
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

  // Whenever we're not dragging, the displayed rect is just the
  // committed model values. While dragging the transient state wins.
  const displayBodyX = transient?.bodyX ?? baseBodyPx.x;
  const displayBodyY = transient?.bodyY ?? baseBodyPx.y;
  const displayWidth = transient?.widthPx ?? image.anchor.widthPx;
  const displayHeight = transient?.heightPx ?? image.anchor.heightPx;

  useEffect(() => {
    if (!dragMode) return;
    const onMove = (e: MouseEvent) => {
      if (dragMode.kind === "move") {
        const dx = e.clientX - dragMode.startMouseX;
        const dy = e.clientY - dragMode.startMouseY;
        setTransient({
          bodyX: Math.max(0, dragMode.startBodyX + dx),
          bodyY: Math.max(0, dragMode.startBodyY + dy),
          widthPx: image.anchor.widthPx,
          heightPx: image.anchor.heightPx,
        });
        return;
      }
      const dx = e.clientX - dragMode.startMouseX;
      const dy = e.clientY - dragMode.startMouseY;
      const next = applyResize(dragMode, dx, dy);
      setTransient(next);
    };
    const onUp = () => {
      const final = transientRef.current;
      setDragMode(null);
      setTransient(null);
      if (!final) return;
      if (dragMode.kind === "move") {
        const anchor = bodyPxToAnchor(final.bodyX, final.bodyY, colXs, rowYs);
        if (
          anchor.fromRow !== image.anchor.fromRow ||
          anchor.fromCol !== image.anchor.fromCol ||
          Math.round(anchor.fromOffsetXPx) !== Math.round(image.anchor.fromOffsetXPx) ||
          Math.round(anchor.fromOffsetYPx) !== Math.round(image.anchor.fromOffsetYPx)
        ) {
          onMoveCommit(anchor);
        }
      } else {
        const widthPx = Math.max(8, Math.round(final.widthPx));
        const heightPx = Math.max(8, Math.round(final.heightPx));
        if (widthPx !== image.anchor.widthPx || heightPx !== image.anchor.heightPx) {
          onResizeCommit({ widthPx, heightPx });
        }
        if (dragMode.handle.includes("w") || dragMode.handle.includes("n")) {
          // Resize from a top/left handle also moves the anchor.
          const anchor = bodyPxToAnchor(final.bodyX, final.bodyY, colXs, rowYs);
          if (
            anchor.fromRow !== image.anchor.fromRow ||
            anchor.fromCol !== image.anchor.fromCol ||
            Math.round(anchor.fromOffsetXPx) !== Math.round(image.anchor.fromOffsetXPx) ||
            Math.round(anchor.fromOffsetYPx) !== Math.round(image.anchor.fromOffsetYPx)
          ) {
            onMoveCommit(anchor);
          }
        }
      }
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
  }, [dragMode, colXs, rowYs, image.anchor, onMoveCommit, onResizeCommit]);

  const containerStyle: CSSProperties = {
    position: "absolute",
    top: headerOffset.y + displayBodyY,
    left: headerOffset.x + displayBodyX,
    width: displayWidth,
    height: displayHeight,
    boxSizing: "border-box",
    cursor: dragMode?.kind === "move" ? "grabbing" : "grab",
    zIndex: selected ? 12 : 10,
    userSelect: "none",
  };

  const overlay: ReactNode = selected ? (
    <>
      <div
        aria-hidden
        style={{
          position: "absolute",
          inset: 0,
          border: "1.5px solid var(--ai-violet, #7c3aed)",
          pointerEvents: "none",
        }}
      />
      {(["nw", "n", "ne", "e", "se", "s", "sw", "w"] as ResizeHandle[]).map((h) => (
        <div
          key={`handle-${h}`}
          data-testid={`image-handle-${imageId}-${h}`}
          aria-label={`Resize ${h}`}
          onMouseDown={(e) => {
            e.preventDefault();
            e.stopPropagation();
            onSelect();
            setTransient({
              bodyX: baseBodyPx.x,
              bodyY: baseBodyPx.y,
              widthPx: image.anchor.widthPx,
              heightPx: image.anchor.heightPx,
            });
            setDragMode({
              kind: "resize",
              handle: h,
              startMouseX: e.clientX,
              startMouseY: e.clientY,
              startBodyX: baseBodyPx.x,
              startBodyY: baseBodyPx.y,
              startWidth: image.anchor.widthPx,
              startHeight: image.anchor.heightPx,
            });
          }}
          style={{
            position: "absolute",
            ...handleStyle(h),
            width: HANDLE_SIZE,
            height: HANDLE_SIZE,
            background: "white",
            border: "1.5px solid var(--ai-violet, #7c3aed)",
            borderRadius: 2,
            cursor: RESIZE_CURSORS[h],
          }}
        />
      ))}
    </>
  ) : null;

  return (
    <div
      data-testid={`image-${imageId}`}
      data-image-id={imageId}
      role="img"
      aria-label={image.altText ?? image.name ?? "Image"}
      style={containerStyle}
      onMouseDown={(e) => {
        e.preventDefault();
        e.stopPropagation();
        onSelect();
        setTransient({
          bodyX: baseBodyPx.x,
          bodyY: baseBodyPx.y,
          widthPx: image.anchor.widthPx,
          heightPx: image.anchor.heightPx,
        });
        setDragMode({
          kind: "move",
          startMouseX: e.clientX,
          startMouseY: e.clientY,
          startBodyX: baseBodyPx.x,
          startBodyY: baseBodyPx.y,
        });
      }}
    >
      {src ? (
        <img
          src={src}
          alt={image.altText ?? image.name ?? ""}
          draggable={false}
          style={{
            display: "block",
            width: "100%",
            height: "100%",
            objectFit: "fill",
            pointerEvents: "none",
          }}
        />
      ) : (
        <div
          aria-hidden
          style={{
            width: "100%",
            height: "100%",
            background: "var(--surface, #f3f4f6)",
            border: "1px dashed var(--divider, #d1d5db)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: 11,
            color: "var(--secondary, #6b7280)",
          }}
        >
          (image unavailable)
        </div>
      )}
      {overlay}
    </div>
  );
}

function applyResize(
  drag: Extract<DragMode, { kind: "resize" }>,
  dx: number,
  dy: number
): { bodyX: number; bodyY: number; widthPx: number; heightPx: number } {
  let { startBodyX: x, startBodyY: y, startWidth: w, startHeight: h } = drag;
  const minSize = 8;
  if (drag.handle.includes("e")) w = Math.max(minSize, drag.startWidth + dx);
  if (drag.handle.includes("s")) h = Math.max(minSize, drag.startHeight + dy);
  if (drag.handle.includes("w")) {
    const newW = Math.max(minSize, drag.startWidth - dx);
    x = drag.startBodyX + (drag.startWidth - newW);
    w = newW;
  }
  if (drag.handle.includes("n")) {
    const newH = Math.max(minSize, drag.startHeight - dy);
    y = drag.startBodyY + (drag.startHeight - newH);
    h = newH;
  }
  return { bodyX: Math.max(0, x), bodyY: Math.max(0, y), widthPx: w, heightPx: h };
}

function handleStyle(h: ResizeHandle): CSSProperties {
  const half = -Math.floor(HANDLE_SIZE / 2);
  switch (h) {
    case "nw":
      return { top: half, left: half };
    case "n":
      return { top: half, left: `calc(50% - ${HANDLE_SIZE / 2}px)` };
    case "ne":
      return { top: half, right: half };
    case "e":
      return { top: `calc(50% - ${HANDLE_SIZE / 2}px)`, right: half };
    case "se":
      return { bottom: half, right: half };
    case "s":
      return { bottom: half, left: `calc(50% - ${HANDLE_SIZE / 2}px)` };
    case "sw":
      return { bottom: half, left: half };
    case "w":
      return { top: `calc(50% - ${HANDLE_SIZE / 2}px)`, left: half };
  }
}
