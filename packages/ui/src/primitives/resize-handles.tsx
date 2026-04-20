"use client";

import { type CSSProperties, type ReactNode } from "react";
import { cn } from "../lib/cn";

/**
 * D8 — shared 8-handle resize chrome.
 *
 * Used by every editor that paints "selected box with corner +
 * edge handles": the XLSX `ImageOverlay` / `ChartOverlay`, the
 * DOCX `ImageResizeOverlay`, and the PPTX `CropOverlay`. The PPTX
 * `SlideCanvas` shape selection chrome stays bespoke — it lives
 * inside an SVG layer with its own hit-zone vs visible-handle
 * split, which doesn't map cleanly onto an HTML-div primitive.
 *
 * Design rationale:
 *   * Controlled-style: the primitive owns nothing except the
 *     handle DOM + cursor mapping. Position math, transient drag
 *     state, snapping, EMU conversion, and command dispatch all
 *     stay in the parent — each editor has its own coordinate
 *     system that's not worth abstracting.
 *   * `onHandleGrab` reports `clientX`/`clientY`/`shiftKey` (plus
 *     which side and whether it's a corner) at mousedown. The
 *     parent installs window-level mousemove/mouseup listeners
 *     itself; the primitive intentionally doesn't, to keep its
 *     contract a single render-time prop.
 *   * Shift-proportional behaviour ALSO stays parent-side because
 *     aspect-ratio rules differ per consumer (image natural ratio
 *     vs chart current ratio vs cropped picture ratio).
 *
 * Visual: 8 absolutely-positioned `<div>`s with white fill +
 * accent-colour outline, sized 8 px by default. The parent
 * MUST be `position: relative` (or otherwise establish a
 * containing block); the primitive does not wrap its handles
 * in any container of its own to keep the DOM flat.
 */

export type ResizeHandleSide =
  | "nw"
  | "n"
  | "ne"
  | "w"
  | "e"
  | "sw"
  | "s"
  | "se";

export const ALL_RESIZE_HANDLE_SIDES: ReadonlyArray<ResizeHandleSide> = [
  "nw",
  "n",
  "ne",
  "e",
  "se",
  "s",
  "sw",
  "w",
];

const CORNER_SIDES: ReadonlySet<ResizeHandleSide> = new Set<ResizeHandleSide>([
  "nw",
  "ne",
  "se",
  "sw",
]);

/**
 * Cursor name to use for each handle. Owned by the primitive so
 * every consumer renders the same cursor for the same side; the
 * parent never needs to thread a per-handle cursor map through.
 */
const RESIZE_CURSORS: Record<ResizeHandleSide, string> = {
  nw: "nwse-resize",
  n: "ns-resize",
  ne: "nesw-resize",
  e: "ew-resize",
  se: "nwse-resize",
  s: "ns-resize",
  sw: "nesw-resize",
  w: "ew-resize",
};

export function isCornerSide(side: ResizeHandleSide): boolean {
  return CORNER_SIDES.has(side);
}

export interface ResizeHandleGrabInfo {
  readonly side: ResizeHandleSide;
  readonly isCorner: boolean;
  readonly shiftKey: boolean;
  readonly clientX: number;
  readonly clientY: number;
}

export interface ResizeHandlesProps {
  /** Which sides to render. Default: all eight. */
  readonly sides?: ReadonlyArray<ResizeHandleSide>;
  /** Visual size of each handle in CSS pixels. Default 8. */
  readonly handleSizePx?: number;
  /**
   * Called on mousedown. The parent installs window-level
   * mousemove/mouseup listeners to compute the running rect. The
   * primitive only tells the parent which handle was grabbed and
   * the starting client coordinates plus the modifier state.
   *
   * The primitive calls `e.preventDefault()` and
   * `e.stopPropagation()` BEFORE invoking this callback so consumers
   * don't have to remember to.
   */
  readonly onHandleGrab: (info: ResizeHandleGrabInfo) => void;
  /** Optional className applied to each handle. Useful for ring colour. */
  readonly handleClassName?: string;
  /**
   * Optional inline style overrides applied last on each handle —
   * lets specific consumers (e.g. the DOCX overlay that lives
   * inside a `pointer-events-none` wrapper) re-enable pointer
   * events or tweak colours without forking the primitive.
   */
  readonly handleStyle?: CSSProperties;
  /**
   * If supplied, each handle gets `data-testid={prefix}-{side}`
   * (e.g. `image-handle-img1-nw`). Without it, only the static
   * `data-side` attribute is set, which is enough for the
   * primitive's own unit tests.
   */
  readonly dataTestIdPrefix?: string;
  /** Accessible label per side. Defaults to `Resize ${side}`. */
  readonly handleLabel?: (side: ResizeHandleSide) => string;
}

/**
 * Renders eight absolutely-positioned drag handles inside the
 * caller's relative-positioned bounding box. See module doc for
 * the contract; see `resize-handles.test.tsx` for the behavioural
 * pin.
 */
export function ResizeHandles(props: ResizeHandlesProps): ReactNode {
  const {
    sides = ALL_RESIZE_HANDLE_SIDES,
    handleSizePx = 8,
    onHandleGrab,
    handleClassName,
    handleStyle,
    dataTestIdPrefix,
    handleLabel,
  } = props;

  return (
    <>
      {sides.map((side) => {
        const isCorner = CORNER_SIDES.has(side);
        const baseStyle = computeHandleStyle(side, handleSizePx);
        return (
          <div
            key={`resize-handle-${side}`}
            role="presentation"
            data-side={side}
            data-corner={isCorner ? "true" : "false"}
            data-testid={
              dataTestIdPrefix ? `${dataTestIdPrefix}-${side}` : undefined
            }
            aria-label={handleLabel ? handleLabel(side) : `Resize ${side}`}
            className={cn(handleClassName)}
            style={{
              ...baseStyle,
              cursor: RESIZE_CURSORS[side],
              ...handleStyle,
            }}
            onMouseDown={(e) => {
              e.preventDefault();
              e.stopPropagation();
              onHandleGrab({
                side,
                isCorner,
                shiftKey: e.shiftKey,
                clientX: e.clientX,
                clientY: e.clientY,
              });
            }}
          />
        );
      })}
    </>
  );
}

function computeHandleStyle(side: ResizeHandleSide, sizePx: number): CSSProperties {
  // Centre the handle on the corner / edge midpoint by translating
  // by half its size on each relevant axis. We use `top`/`left`
  // (and a calc for midpoints) rather than `transform` so the
  // handle stays pixel-snapped at integer zooms.
  const half = -Math.floor(sizePx / 2);
  const mid = `calc(50% - ${sizePx / 2}px)`;
  const base: CSSProperties = {
    position: "absolute",
    width: sizePx,
    height: sizePx,
    boxSizing: "border-box",
    background: "white",
    border: "1.5px solid currentColor",
    borderRadius: 2,
    pointerEvents: "auto",
    touchAction: "none",
    userSelect: "none",
  };
  switch (side) {
    case "nw":
      return { ...base, top: half, left: half };
    case "n":
      return { ...base, top: half, left: mid };
    case "ne":
      return { ...base, top: half, right: half };
    case "e":
      return { ...base, top: mid, right: half };
    case "se":
      return { ...base, bottom: half, right: half };
    case "s":
      return { ...base, bottom: half, left: mid };
    case "sw":
      return { ...base, bottom: half, left: half };
    case "w":
      return { ...base, top: mid, left: half };
  }
}
