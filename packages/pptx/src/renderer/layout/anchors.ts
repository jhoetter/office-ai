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
 */
export function anchorsFor(shapeId: string, box: BoundingBox): ShapeAnchor[] {
  const cx = Math.round(box.x + box.cx / 2);
  const cy = Math.round(box.y + box.cy / 2);
  const ts: ReadonlyArray<number> = [0.25, 0.5, 0.75];
  const out: ShapeAnchor[] = [];
  for (const t of ts) {
    out.push({ shapeId, side: "n", x: Math.round(box.x + box.cx * t), y: box.y, t });
    out.push({ shapeId, side: "s", x: Math.round(box.x + box.cx * t), y: box.y + box.cy, t });
    out.push({ shapeId, side: "w", x: box.x, y: Math.round(box.y + box.cy * t), t });
    out.push({ shapeId, side: "e", x: box.x + box.cx, y: Math.round(box.y + box.cy * t), t });
  }
  out.push({ shapeId, side: "center", x: cx, y: cy, t: 0.5 });
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
 */
export function snapToAnchor(
  point: { x: number; y: number },
  shapes: ReadonlyArray<{ id: string; box: BoundingBox }>,
  thresholdEmu: number
): AnchorSnapResult {
  let best: ShapeAnchor | null = null;
  let bestDist = Infinity;
  const nearby: ShapeAnchor[] = [];
  for (const s of shapes) {
    for (const a of anchorsFor(s.id, s.box)) {
      const dx = a.x - point.x;
      const dy = a.y - point.y;
      const d = Math.hypot(dx, dy);
      if (d > thresholdEmu) continue;
      nearby.push(a);
      if (d < bestDist) {
        best = a;
        bestDist = d;
      }
    }
  }
  if (!best) return { dx: 0, dy: 0, anchor: null, nearby };
  return { dx: best.x - point.x, dy: best.y - point.y, anchor: best, nearby };
}
