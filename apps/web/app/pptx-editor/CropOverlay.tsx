"use client";

import * as React from "react";
import { ResizeHandles, type ResizeHandleSide } from "@officeai/ui";
import type { Picture } from "@officeai/pptx";

/**
 * D7 — image-crop overlay rendered on top of the slide canvas while
 * the user is in crop mode. The user sees their picture's bounding
 * box dimmed everywhere except the inner "kept" rectangle, can drag
 * any of 8 handles to resize the rect, and commits with Enter or
 * an outside-click (Escape cancels).
 *
 * Coordinate system:
 *   • The crop rectangle is tracked as percentages [0..1] of the
 *     picture's displayed bounding box. We deliberately keep this
 *     overlay agnostic of EMU and the slide's intrinsic image size
 *     — the host translates the percentages into a `pptx:crop-picture`
 *     payload (multiplied by 100) when committing.
 *   • Pixel positions are derived per-render from the slide-card
 *     `<div>` (rendered by `SlideCanvas`) via a ResizeObserver, so
 *     the overlay re-aligns when the surface resizes / zooms.
 *
 * Commit semantics: the percentages map directly to OOXML
 * `<a:srcRect>` sides. If the user dragged the rect to e.g.
 * `{ left: 0.05, top: 0.10, right: 0.85, bottom: 0.95 }`, the
 * crop command receives `leftPct=5, topPct=10, rightPct=15,
 * bottomPct=5`. The handler validates that `leftPct + rightPct < 100`
 * and `topPct + bottomPct < 100`; we enforce a minimum 1 % gap on
 * each axis here to keep the overlay from accidentally producing a
 * rejected commit.
 */

export interface CropCommitPayload {
  readonly leftPct: number;
  readonly topPct: number;
  readonly rightPct: number;
  readonly bottomPct: number;
}

export interface CropOverlayProps {
  /** Picture being cropped — drives the displayed bounding box. */
  readonly picture: Picture;
  /** Slide size in EMU. Needed to convert picture EMU → slide-relative %. */
  readonly slideSize: { readonly cxEmu: number; readonly cyEmu: number };
  /** Container the overlay is mounted in (the same box that wraps `SlideCanvas`). */
  readonly containerRef: React.RefObject<HTMLElement | null>;
  readonly onCommit: (crop: CropCommitPayload) => void;
  readonly onCancel: () => void;
}

type Side = ResizeHandleSide;

/** Crop rect, expressed as fractions [0, 1] of the picture's bounding box. */
interface CropFracs {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

interface DragState {
  readonly side: Side;
  readonly startX: number;
  readonly startY: number;
  readonly startCrop: CropFracs;
  /** Picture rect (px) captured at gesture start. */
  readonly pictureRect: { x: number; y: number; w: number; h: number };
  readonly cleanup: () => void;
}

const MIN_FRAC = 0.01;

export function CropOverlay(props: CropOverlayProps): React.ReactElement | null {
  const { picture, slideSize, containerRef, onCommit, onCancel } = props;
  const [tick, setTick] = React.useState(0);
  // Crop fractions inside the picture's bounding box. Default to
  // "no crop" (the full picture). We deliberately don't seed from
  // `picture.srcRect` — this overlay's mental model is "draw a new
  // crop rect on the visible image", not "edit the existing crop".
  const [crop, setCrop] = React.useState<CropFracs>({ left: 0, top: 0, right: 1, bottom: 1 });
  const cropRef = React.useRef(crop);
  cropRef.current = crop;
  const dragRef = React.useRef<DragState | null>(null);
  const overlayRef = React.useRef<HTMLDivElement | null>(null);

  // Trigger a re-render whenever the surrounding container or the
  // slide card resizes — that's the only signal we need to recompute
  // pixel positions, since the picture's slide-EMU geometry is
  // immutable for the duration of the crop session.
  React.useEffect(() => {
    const host = containerRef.current;
    if (!host) return;
    const ros: ResizeObserver[] = [];
    const bump = () => setTick((t) => t + 1);
    const o1 = new ResizeObserver(bump);
    o1.observe(host);
    ros.push(o1);
    const card = host.querySelector<HTMLDivElement>("[data-testid='pptx-slide-card']");
    if (card) {
      const o2 = new ResizeObserver(bump);
      o2.observe(card);
      ros.push(o2);
    }
    return () => ros.forEach((r) => r.disconnect());
  }, [containerRef]);

  // Outside-click commit + Esc cancel + Enter commit.
  React.useEffect(() => {
    const onPointerDown = (e: PointerEvent) => {
      const target = e.target as Node | null;
      const root = overlayRef.current;
      if (!root || !target) return;
      if (root.contains(target)) return;
      // Outside the overlay → commit and exit.
      onCommit(toPayload(cropRef.current));
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        onCancel();
        return;
      }
      if (e.key === "Enter") {
        e.preventDefault();
        onCommit(toPayload(cropRef.current));
      }
    };
    // Use capture for pointerdown so we beat the canvas's own pointer
    // handler — otherwise outside-clicks would land on a shape and
    // mutate selection before we get a chance to commit.
    window.addEventListener("pointerdown", onPointerDown, true);
    window.addEventListener("keydown", onKey);
    return () => {
      window.removeEventListener("pointerdown", onPointerDown, true);
      window.removeEventListener("keydown", onKey);
    };
  }, [onCommit, onCancel]);

  void tick;
  const host = containerRef.current;
  const pictureRect = host ? computePictureRect(host, picture, slideSize) : null;
  if (!pictureRect) return null;

  const cropPx = {
    left: pictureRect.x + crop.left * pictureRect.w,
    top: pictureRect.y + crop.top * pictureRect.h,
    right: pictureRect.x + crop.right * pictureRect.w,
    bottom: pictureRect.y + crop.bottom * pictureRect.h,
  };

  const beginDrag = (side: Side, startX: number, startY: number) => {
    // Window-level mouse listeners replace the previous pointer-
    // capture trick: the shared `ResizeHandles` primitive emits
    // mousedown only, and we still need to keep tracking moves
    // even when the cursor leaves the small handle square.
    const onMove = (ev: MouseEvent) => {
      const drag = dragRef.current;
      if (!drag) return;
      const dx = (ev.clientX - drag.startX) / drag.pictureRect.w;
      const dy = (ev.clientY - drag.startY) / drag.pictureRect.h;
      setCrop((prev) => applyDrag(drag.startCrop, drag.side, dx, dy) ?? prev);
    };
    const onUp = () => {
      const drag = dragRef.current;
      if (!drag) return;
      drag.cleanup();
      dragRef.current = null;
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    dragRef.current = {
      side,
      startX,
      startY,
      startCrop: { ...cropRef.current },
      pictureRect,
      cleanup: () => {
        window.removeEventListener("mousemove", onMove);
        window.removeEventListener("mouseup", onUp);
      },
    };
  };

  return (
    <div
      ref={overlayRef}
      data-testid="pptx-crop-overlay"
      style={{
        position: "absolute",
        inset: 0,
        pointerEvents: "none",
        zIndex: 25,
      }}
    >
      {/* Four dim strips around the crop rect (no clip-path so the
          overlay renders identically across browsers / Safari). Each
          strip absorbs pointer events so the user can't accidentally
          drag the underlying shape while in crop mode. */}
      <DimStrip
        x={pictureRect.x}
        y={pictureRect.y}
        w={pictureRect.w}
        h={cropPx.top - pictureRect.y}
      />
      <DimStrip
        x={pictureRect.x}
        y={cropPx.bottom}
        w={pictureRect.w}
        h={pictureRect.y + pictureRect.h - cropPx.bottom}
      />
      <DimStrip
        x={pictureRect.x}
        y={cropPx.top}
        w={cropPx.left - pictureRect.x}
        h={cropPx.bottom - cropPx.top}
      />
      <DimStrip
        x={cropPx.right}
        y={cropPx.top}
        w={pictureRect.x + pictureRect.w - cropPx.right}
        h={cropPx.bottom - cropPx.top}
      />
      {/* Crop rect outline + 8 drag handles. The handles live INSIDE
          this rect so the shared `ResizeHandles` primitive can
          position them at corner / midpoint of the rect's
          relative box. The outline div itself is `pointer-events:
          none` (only the dim strips and handles absorb clicks)
          so a click in the kept-area still falls through to the
          outside-click commit handler. */}
      <div
        className="text-zinc-800"
        style={{
          position: "absolute",
          left: cropPx.left,
          top: cropPx.top,
          width: cropPx.right - cropPx.left,
          height: cropPx.bottom - cropPx.top,
          border: "1px solid white",
          boxShadow: "0 0 0 1px rgba(0,0,0,0.5)",
          pointerEvents: "none",
        }}
      >
        <ResizeHandles
          handleSizePx={10}
          dataTestIdPrefix="pptx-crop-handle"
          handleStyle={{
            // Match the original chrome (square corners, dark grey
            // border) and re-enable pointer events because the
            // outline parent disables them.
            borderRadius: 0,
            borderWidth: 1,
            pointerEvents: "auto",
          }}
          onHandleGrab={(info) => {
            beginDrag(info.side, info.clientX, info.clientY);
          }}
        />
      </div>
    </div>
  );
}

function DimStrip({ x, y, w, h }: { x: number; y: number; w: number; h: number }): React.ReactElement | null {
  if (w <= 0 || h <= 0) return null;
  return (
    <div
      style={{
        position: "absolute",
        left: x,
        top: y,
        width: w,
        height: h,
        background: "rgba(0,0,0,0.4)",
        pointerEvents: "auto",
      }}
    />
  );
}

function computePictureRect(
  host: HTMLElement,
  picture: Picture,
  slideSize: { cxEmu: number; cyEmu: number }
): { x: number; y: number; w: number; h: number } | null {
  if (!picture.position || !picture.size) return null;
  const card = host.querySelector<HTMLDivElement>("[data-testid='pptx-slide-card']");
  if (!card) return null;
  const hostRect = host.getBoundingClientRect();
  const cardRect = card.getBoundingClientRect();
  if (cardRect.width === 0 || cardRect.height === 0) return null;
  const sx = cardRect.width / slideSize.cxEmu;
  const sy = cardRect.height / slideSize.cyEmu;
  const x = cardRect.left - hostRect.left + picture.position.xEmu * sx;
  const y = cardRect.top - hostRect.top + picture.position.yEmu * sy;
  const w = picture.size.cxEmu * sx;
  const h = picture.size.cyEmu * sy;
  return { x, y, w, h };
}

/**
 * Apply a drag delta (in fractions of the picture box) to one side
 * of the crop rect. Returns `null` when the drag would collapse the
 * rect below `MIN_FRAC` on either axis (so we don't ratchet down to
 * zero). Each side is clamped into `[0, 1]` against the opposite
 * side's current position.
 */
function applyDrag(start: CropFracs, side: Side, dx: number, dy: number): CropFracs | null {
  const next: CropFracs = { ...start };
  const clamp = (v: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, v));
  if (side === "w" || side === "nw" || side === "sw") {
    next.left = clamp(start.left + dx, 0, start.right - MIN_FRAC);
  }
  if (side === "e" || side === "ne" || side === "se") {
    next.right = clamp(start.right + dx, start.left + MIN_FRAC, 1);
  }
  if (side === "n" || side === "nw" || side === "ne") {
    next.top = clamp(start.top + dy, 0, start.bottom - MIN_FRAC);
  }
  if (side === "s" || side === "sw" || side === "se") {
    next.bottom = clamp(start.bottom + dy, start.top + MIN_FRAC, 1);
  }
  return next;
}

function toPayload(crop: CropFracs): CropCommitPayload {
  // Crop rect is "what to keep"; OOXML `<a:srcRect>` measures
  // "what to chop off from each side", so right/bottom invert.
  const round = (n: number) => Math.round(n * 1000) / 10; // 0.1 % precision
  return {
    leftPct: round(crop.left * 100),
    topPct: round(crop.top * 100),
    rightPct: round((1 - crop.right) * 100),
    bottomPct: round((1 - crop.bottom) * 100),
  };
}
