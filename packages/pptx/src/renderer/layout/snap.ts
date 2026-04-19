import type { SlideSize } from "../../model/types.js";
import type { BoundingBox } from "./shape.js";

/**
 * A single alignment guide, rendered as a thin coloured line by the
 * canvas while the user drags. Guides are axis-aligned: a `vertical`
 * guide is a vertical line at `value` (an x coordinate) and a
 * `horizontal` guide is a horizontal line at `value` (a y coordinate).
 *
 * `spanStart` / `spanEnd` describe the orthogonal extent the line
 * should cover so the user sees exactly which two shapes it's tying
 * together (Figma-style: from the topmost edge involved to the
 * bottommost). For slide-edge guides the span is the full slide.
 */
export interface SnapGuide {
  readonly axis: "vertical" | "horizontal";
  readonly value: number;
  readonly spanStart: number;
  readonly spanEnd: number;
  /**
   * What this guide is matching: an edge (start/end), the centre, or
   * the slide chrome (slide bounds + slide centre). Used by the
   * renderer to colour-code centre vs edge guides if it likes.
   */
  readonly kind: "edge" | "center" | "slide";
}

export interface SnapResult {
  /** Delta to add to every dragged shape's x. Zero when no snap fired. */
  readonly snapDx: number;
  /** Delta to add to every dragged shape's y. Zero when no snap fired. */
  readonly snapDy: number;
  readonly guides: ReadonlyArray<SnapGuide>;
}

/**
 * Compute snap deltas + guide lines for a moving shape (or the union
 * box of a multi-shape drag) given the bounding boxes of every other
 * shape on the slide. We pick the SINGLE best snap per axis (closest
 * within `thresholdEmu`); the other near-miss candidates within the
 * same threshold are also returned as guides so the user sees every
 * alignment, but they don't influence the snap delta.
 *
 * Snap candidates per axis are: left edge, centre, right edge of
 * every other shape, plus the slide's left/centre/right edges. Same
 * for the y axis.
 */
export function computeSnap(
  movingBox: BoundingBox,
  others: ReadonlyArray<{ id: string; box: BoundingBox }>,
  slideSize: SlideSize,
  thresholdEmu: number
): SnapResult {
  const xCands = collectXCandidates(others, slideSize);
  const yCands = collectYCandidates(others, slideSize);

  const xSnap = pickBestSnap(
    movingBox.x,
    movingBox.x + movingBox.cx / 2,
    movingBox.x + movingBox.cx,
    xCands,
    thresholdEmu
  );
  const ySnap = pickBestSnap(
    movingBox.y,
    movingBox.y + movingBox.cy / 2,
    movingBox.y + movingBox.cy,
    yCands,
    thresholdEmu
  );

  const snapDx = xSnap?.delta ?? 0;
  const snapDy = ySnap?.delta ?? 0;

  // Compose guides: one per axis for the picked snap, plus near-miss
  // edges so the user sees every alignment opportunity. We intentionally
  // include only the guides whose target position equals the SNAPPED
  // moving edge — otherwise we'd render a forest of unrelated lines.
  const finalLeft = movingBox.x + snapDx;
  const finalRight = finalLeft + movingBox.cx;
  const finalCenterX = finalLeft + movingBox.cx / 2;
  const finalTop = movingBox.y + snapDy;
  const finalBottom = finalTop + movingBox.cy;
  const finalCenterY = finalTop + movingBox.cy / 2;

  const guides: SnapGuide[] = [];
  if (xSnap) {
    for (const c of xCands) {
      if (c.value !== xSnap.target) continue;
      if (
        nearlyEqual(finalLeft, c.value) ||
        nearlyEqual(finalCenterX, c.value) ||
        nearlyEqual(finalRight, c.value)
      ) {
        guides.push(makeVerticalGuide(c, movingBox, snapDx, snapDy));
      }
    }
  }
  if (ySnap) {
    for (const c of yCands) {
      if (c.value !== ySnap.target) continue;
      if (
        nearlyEqual(finalTop, c.value) ||
        nearlyEqual(finalCenterY, c.value) ||
        nearlyEqual(finalBottom, c.value)
      ) {
        guides.push(makeHorizontalGuide(c, movingBox, snapDx, snapDy));
      }
    }
  }
  return { snapDx, snapDy, guides };
}

interface SnapCandidate {
  /** Position along the axis. */
  readonly value: number;
  readonly kind: SnapGuide["kind"];
  /** For shape-derived guides: the involved shape's box (used to size span). */
  readonly refBox: BoundingBox | null;
}

function collectXCandidates(
  others: ReadonlyArray<{ id: string; box: BoundingBox }>,
  slideSize: SlideSize
): SnapCandidate[] {
  const out: SnapCandidate[] = [
    { value: 0, kind: "slide", refBox: null },
    { value: Math.round(slideSize.cxEmu / 2), kind: "slide", refBox: null },
    { value: slideSize.cxEmu, kind: "slide", refBox: null },
  ];
  for (const o of others) {
    out.push({ value: o.box.x, kind: "edge", refBox: o.box });
    out.push({ value: Math.round(o.box.x + o.box.cx / 2), kind: "center", refBox: o.box });
    out.push({ value: o.box.x + o.box.cx, kind: "edge", refBox: o.box });
  }
  return out;
}

function collectYCandidates(
  others: ReadonlyArray<{ id: string; box: BoundingBox }>,
  slideSize: SlideSize
): SnapCandidate[] {
  const out: SnapCandidate[] = [
    { value: 0, kind: "slide", refBox: null },
    { value: Math.round(slideSize.cyEmu / 2), kind: "slide", refBox: null },
    { value: slideSize.cyEmu, kind: "slide", refBox: null },
  ];
  for (const o of others) {
    out.push({ value: o.box.y, kind: "edge", refBox: o.box });
    out.push({ value: Math.round(o.box.y + o.box.cy / 2), kind: "center", refBox: o.box });
    out.push({ value: o.box.y + o.box.cy, kind: "edge", refBox: o.box });
  }
  return out;
}

interface PickedSnap {
  readonly target: number;
  readonly delta: number;
}

/**
 * Pick the closest snap candidate for the given three "moving edges"
 * (left/centre/right or top/centre/bottom). Each edge can match a
 * candidate; we pick the (edge, candidate) pair with the smallest
 * absolute delta. When two pairs tie at exactly the same distance, we
 * prefer matching the moving CENTRE (gives a more "magnetic" feel).
 */
function pickBestSnap(
  edgeStart: number,
  edgeCenter: number,
  edgeEnd: number,
  candidates: ReadonlyArray<SnapCandidate>,
  thresholdEmu: number
): PickedSnap | null {
  let best: { delta: number; abs: number; centerWeight: number } | null = null;
  for (const c of candidates) {
    const tries: ReadonlyArray<{ delta: number; centerWeight: number }> = [
      { delta: c.value - edgeStart, centerWeight: 0 },
      { delta: c.value - edgeCenter, centerWeight: 1 },
      { delta: c.value - edgeEnd, centerWeight: 0 },
    ];
    for (const t of tries) {
      const abs = Math.abs(t.delta);
      if (abs > thresholdEmu) continue;
      if (best === null || abs < best.abs || (abs === best.abs && t.centerWeight > best.centerWeight)) {
        best = { delta: t.delta, abs, centerWeight: t.centerWeight };
      }
    }
  }
  if (!best) return null;
  // Reverse-engineer which candidate value the chosen edge maps to.
  // Easier: target is whichever moving edge + delta lands on; for the
  // caller we only need the target value (for guide selection).
  // We pick "the value the picked candidate yielded" — store via delta+edge
  const target =
    best.centerWeight === 1 ? edgeCenter + best.delta : pickEdgeForDelta(edgeStart, edgeEnd, best.delta);
  return { target, delta: best.delta };
}

function pickEdgeForDelta(edgeStart: number, edgeEnd: number, delta: number): number {
  // We don't actually know which of start/end produced this delta, so
  // recompute the two possibilities and pick whichever is integer-equal
  // to a sensible value. In practice only one edge produces the abs
  // minimum at a given threshold so either way we end up at the same
  // target; we just pick the closer one.
  const a = edgeStart + delta;
  const b = edgeEnd + delta;
  return Math.abs(a - Math.round(a)) <= Math.abs(b - Math.round(b)) ? a : b;
}

function nearlyEqual(a: number, b: number): boolean {
  return Math.abs(a - b) < 1; // sub-EMU tolerance is enough; avoids float noise
}

function makeVerticalGuide(c: SnapCandidate, movingBox: BoundingBox, dx: number, dy: number): SnapGuide {
  const movedBox = {
    x: movingBox.x + dx,
    y: movingBox.y + dy,
    cx: movingBox.cx,
    cy: movingBox.cy,
  };
  if (!c.refBox) {
    return {
      axis: "vertical",
      value: c.value,
      spanStart: 0,
      // Slide-derived guides span the full canvas (caller fills in slide cy).
      // We use a sentinel `Infinity` and let the renderer clamp.
      spanEnd: Number.POSITIVE_INFINITY,
      kind: "slide",
    };
  }
  const top = Math.min(movedBox.y, c.refBox.y);
  const bottom = Math.max(movedBox.y + movedBox.cy, c.refBox.y + c.refBox.cy);
  return { axis: "vertical", value: c.value, spanStart: top, spanEnd: bottom, kind: c.kind };
}

function makeHorizontalGuide(c: SnapCandidate, movingBox: BoundingBox, dx: number, dy: number): SnapGuide {
  const movedBox = {
    x: movingBox.x + dx,
    y: movingBox.y + dy,
    cx: movingBox.cx,
    cy: movingBox.cy,
  };
  if (!c.refBox) {
    return {
      axis: "horizontal",
      value: c.value,
      spanStart: 0,
      spanEnd: Number.POSITIVE_INFINITY,
      kind: "slide",
    };
  }
  const left = Math.min(movedBox.x, c.refBox.x);
  const right = Math.max(movedBox.x + movedBox.cx, c.refBox.x + c.refBox.cx);
  return { axis: "horizontal", value: c.value, spanStart: left, spanEnd: right, kind: c.kind };
}
