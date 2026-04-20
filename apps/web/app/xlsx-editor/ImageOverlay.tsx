"use client";

import { useEffect, useRef, useState, type CSSProperties, type ReactNode } from "react";
import type { SheetImage } from "@officeai/xlsx";
import { ResizeHandles, type ResizeHandleSide } from "@officeai/ui";
import type { AxisLookup } from "./gridDimensions";
import { useTranslator } from "@/lib/i18n";

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
 * Convert a `(fromRow, fromCol, fromOffsetX, fromOffsetY)` anchor
 * into absolute body-coordinate pixels (i.e. relative to the
 * top-left of the cell area, EXCLUDING the row/column header bands).
 *
 * Accepts either a {@link SheetImage} (legacy callers) or a bare
 * `{ anchor }` object so that {@link SheetChart} — which has the
 * same shape but doesn't carry a `mediaRef` — can reuse the helper
 * without a fake image cast.
 */
export function anchorToBodyPx(
  source: { readonly anchor: { readonly fromRow: number; readonly fromCol: number; readonly fromOffsetXPx: number; readonly fromOffsetYPx: number } },
  colXs: AxisLike,
  rowYs: AxisLike
): { x: number; y: number } {
  const c = clamp(source.anchor.fromCol, 0, colXs.length - 2);
  const r = clamp(source.anchor.fromRow, 0, rowYs.length - 2);
  return {
    x: (colXs[c] ?? 0) + source.anchor.fromOffsetXPx,
    y: (rowYs[r] ?? 0) + source.anchor.fromOffsetYPx,
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

type DragMode =
  | {
      kind: "move";
      startMouseX: number;
      startMouseY: number;
      startBodyX: number;
      startBodyY: number;
    }
  | {
      kind: "resize";
      handle: ResizeHandle;
      shiftKey: boolean;
      aspect: number;
      startMouseX: number;
      startMouseY: number;
      startBodyX: number;
      startBodyY: number;
      startWidth: number;
      startHeight: number;
    };

export type ResizeHandle = ResizeHandleSide;

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
  const { t } = useTranslator();

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
      // Shift-while-dragging-a-corner toggles aspect lock live —
      // releasing Shift mid-drag must let the rect freely deform
      // again, so we re-read the modifier on every move rather
      // than freezing it at grab time.
      const next = applyResize(
        dragMode,
        dx,
        dy,
        e.shiftKey && isCorner(dragMode.handle),
      );
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
      {/* Wrapper colour-channel: ResizeHandles uses `currentColor`
          for its border so any consumer styles the chrome by setting
          `color` on a parent. */}
      <div style={{ position: "absolute", inset: 0, pointerEvents: "none", color: "var(--ai-violet, #7c3aed)" }}>
        <ResizeHandles
          handleSizePx={9}
          dataTestIdPrefix={`image-handle-${imageId}`}
          handleLabel={(side) => t("common.resizeHandle", { handle: side })}
          onHandleGrab={(info) => {
            onSelect();
            setTransient({
              bodyX: baseBodyPx.x,
              bodyY: baseBodyPx.y,
              widthPx: image.anchor.widthPx,
              heightPx: image.anchor.heightPx,
            });
            setDragMode({
              kind: "resize",
              handle: info.side,
              shiftKey: info.shiftKey,
              aspect:
                image.anchor.heightPx > 0
                  ? image.anchor.widthPx / image.anchor.heightPx
                  : 1,
              startMouseX: info.clientX,
              startMouseY: info.clientY,
              startBodyX: baseBodyPx.x,
              startBodyY: baseBodyPx.y,
              startWidth: image.anchor.widthPx,
              startHeight: image.anchor.heightPx,
            });
          }}
        />
      </div>
    </>
  ) : null;

  return (
    <div
      data-testid={`image-${imageId}`}
      data-image-id={imageId}
      role="img"
      aria-label={image.altText ?? image.name ?? t("xlsx.image.imageFallback")}
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

function isCorner(h: ResizeHandle): boolean {
  return h === "nw" || h === "ne" || h === "se" || h === "sw";
}

function applyResize(
  drag: Extract<DragMode, { kind: "resize" }>,
  dx: number,
  dy: number,
  proportional: boolean
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
  if (proportional && isCorner(drag.handle) && drag.aspect > 0) {
    // Whichever axis "moved more" relative to the start size wins;
    // the other axis snaps to the locked aspect ratio. We re-derive
    // x / y from the locked dimensions when an n/w handle is in play
    // so the opposite corner stays pinned (matches PowerPoint /
    // Figma corner-resize behaviour).
    const rW = w / drag.startWidth;
    const rH = h / drag.startHeight;
    const widthDriven = Math.abs(rW - 1) >= Math.abs(rH - 1);
    if (widthDriven) {
      h = Math.max(minSize, w / drag.aspect);
    } else {
      w = Math.max(minSize, h * drag.aspect);
    }
    if (drag.handle.includes("w")) x = drag.startBodyX + (drag.startWidth - w);
    if (drag.handle.includes("n")) y = drag.startBodyY + (drag.startHeight - h);
  }
  return { bodyX: Math.max(0, x), bodyY: Math.max(0, y), widthPx: w, heightPx: h };
}
