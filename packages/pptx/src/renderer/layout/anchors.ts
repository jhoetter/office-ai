import { rotateAroundCenter } from "../../model/connector-geometry.js";
import type { BoundingBox } from "./shape.js";

/**
 * A connection anchor — a fixed, semantically-meaningful point on a
 * shape that line/arrow endpoints can latch onto. We expose four
 * cardinal edge-midpoints plus the centre (good enough for ~95% of
 * real-world diagrams; corners are intentionally NOT exposed because
 * they almost never produce a clean connection in PowerPoint).
 */
export type AnchorSide = "n" | "s" | "e" | "w" | "center";

export interface ShapeAnchor {
  readonly shapeId: string;
  readonly side: AnchorSide;
  readonly x: number;
  readonly y: number;
  /**
   * Position along the side in [0, 1]. 0.5 is the cardinal midpoint
   * (the only legal value for `center`); 0.25 / 0.75 are the
   * quarter-points the canvas exposes alongside the midpoint so users
   * can latch arrows onto edge thirds without leaving the magnetic
   * snap. Persisted onto anchored endpoints as `t` and round-tripped
   * by collapsing back to the nearest cardinal index on save.
   */
  readonly t: number;
}

/**
 * Compute the anchor points for a single shape's bounding box. Each
 * cardinal edge is exposed at three points (t = 0.25, 0.5, 0.75) so
 * connectors can land on edge thirds; `center` remains a single
 * anchor at the bbox centre.
 *
 * `rotationDeg` (default 0) rotates every emitted anchor point
 * clockwise about the bbox centre — same pivot as `wrapWithRotation`
 * in `renderer/svg/shapes.ts` — so connectors snap to the shape's
 * visually rendered edges, not to its underlying axis-aligned bbox.
 * The `side` label is preserved through rotation (a "north" anchor
 * stays the north of the *shape*, even if it's now visually below
 * the centre after a 180° rotation), which keeps round-tripping to
 * the OOXML `cxnLst` index straightforward.
 */
export function anchorsFor(shapeId: string, box: BoundingBox, rotationDeg: number = 0): ShapeAnchor[] {
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy / 2;
  const center = { x: cx, y: cy };
  const ts: ReadonlyArray<number> = [0.25, 0.5, 0.75];
  const out: ShapeAnchor[] = [];
  const emit = (side: ShapeAnchor["side"], x: number, y: number, t: number): void => {
    const r = rotationDeg === 0 ? { x, y } : rotateAroundCenter({ x, y }, center, rotationDeg);
    out.push({ shapeId, side, x: Math.round(r.x), y: Math.round(r.y), t });
  };
  for (const t of ts) {
    emit("n", box.x + box.cx * t, box.y, t);
    emit("s", box.x + box.cx * t, box.y + box.cy, t);
    emit("w", box.x, box.y + box.cy * t, t);
    emit("e", box.x + box.cx, box.y + box.cy * t, t);
  }
  // Centre is invariant under rotation about itself.
  out.push({ shapeId, side: "center", x: Math.round(cx), y: Math.round(cy), t: 0.5 });
  return out;
}

export interface AnchorSnapResult {
  /** Dx to add to the dragged endpoint so it lands on the anchor. */
  readonly dx: number;
  /** Dy to add to the dragged endpoint so it lands on the anchor. */
  readonly dy: number;
  /** The anchor that won, or null if no anchor was in range. */
  readonly anchor: ShapeAnchor | null;
  /**
   * Every anchor within the threshold (so the canvas can show a faded
   * dot on each candidate, plus a solid dot on the winner). Useful UX
   * cue — without it the user can't tell what the snap will do until
   * after it fires.
   */
  readonly nearby: ReadonlyArray<ShapeAnchor>;
}

/**
 * Pick the closest anchor across the given shapes that's within the
 * threshold (Euclidean distance in EMU). Used while dragging a line
 * endpoint to snap it onto another shape.
 *
 * Edge anchors (n/s/e/w) take strict priority over the centre anchor:
 * if any edge candidate is in range we never return centre, even when
 * centre is geometrically closer. This matches PowerPoint/Slides where
 * dragging towards a shape body lands on the nearest edge, not the
 * centroid — landing on the centroid produces routes that visibly cut
 * through the shape and the user almost never wants that.
 */
export function snapToAnchor(
  point: { x: number; y: number },
  shapes: ReadonlyArray<{ id: string; box: BoundingBox; rotation?: number }>,
  thresholdEmu: number
): AnchorSnapResult {
  let bestEdge: ShapeAnchor | null = null;
  let bestEdgeDist = Infinity;
  let bestCenter: ShapeAnchor | null = null;
  let bestCenterDist = Infinity;
  const nearbyEdges: ShapeAnchor[] = [];
  const nearbyCenters: ShapeAnchor[] = [];
  for (const s of shapes) {
    for (const a of anchorsFor(s.id, s.box, s.rotation ?? 0)) {
      const dx = a.x - point.x;
      const dy = a.y - point.y;
      const d = Math.hypot(dx, dy);
      if (d > thresholdEmu) continue;
      if (a.side === "center") {
        nearbyCenters.push(a);
        if (d < bestCenterDist) {
          bestCenter = a;
          bestCenterDist = d;
        }
      } else {
        nearbyEdges.push(a);
        if (d < bestEdgeDist) {
          bestEdge = a;
          bestEdgeDist = d;
        }
      }
    }
  }
  // Edges win whenever any edge anchor is in range. Centre is a
  // fallback used only when the cursor is deep inside a shape with
  // no edge in snap range (rare unless the threshold is very small
  // relative to the shape).
  const best = bestEdge ?? bestCenter;
  // Surface the edge candidates as the "live snap dots" the canvas
  // paints; suppress centre dots entirely so the visible affordance
  // matches the snap behaviour.
  const nearby = bestEdge ? nearbyEdges : [...nearbyEdges, ...nearbyCenters];
  if (!best) return { dx: 0, dy: 0, anchor: null, nearby };
  return { dx: best.x - point.x, dy: best.y - point.y, anchor: best, nearby };
}
