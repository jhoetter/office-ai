import type { CommandHandler, NodeId } from "@officeai/core";
import type { PptxSnapshot, Shape, Slide } from "../model/types.js";
import { shapeBoundingBox, type BoundingBox } from "../renderer/layout/shape.js";
import {
  buildDiff,
  evolveSnapshot,
  findShapeInSlide,
  findSlide,
  makeError,
  replaceShape,
  withSlide,
} from "./helpers.js";
import type { AlignShapesPayload, DistributeShapesPayload } from "./payloads.js";

/**
 * Align two-or-more shapes by edge or centre. Mirrors the PowerPoint
 * "Align" menu (Left/Center/Right/Top/Middle/Bottom). For all six modes
 * we operate on the union bounding box of the selection — left snaps
 * every shape to the union's leftmost x, center-h snaps every shape's
 * x-center to the union's x-center, etc. Shapes without an explicit
 * position (placeholders inheriting from layout) are skipped. Opaque
 * shapes are also skipped because we can't safely re-emit their xfrm.
 */
export const alignShapesHandler: CommandHandler<AlignShapesPayload, PptxSnapshot> = {
  type: "pptx:align-shapes",
  apply(snapshot, payload) {
    if (!payload.shapeIds || payload.shapeIds.length < 2) {
      throw makeError("invalid-payload", "align-shapes requires at least 2 shape ids");
    }
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const targets = collectAlignTargets(slide, payload.shapeIds);
    if (targets.length < 2) {
      throw makeError("invalid-payload", "align-shapes needs at least 2 alignable shapes");
    }
    const union = unionBox(targets.map((t) => t.box));
    const updates = targets.map((t) => alignBox(t.box, union, payload.mode));

    let root = snapshot.root;
    for (let i = 0; i < targets.length; i++) {
      const target = targets[i];
      const next = updates[i];
      if (next.x === target.box.x && next.y === target.box.y) continue;
      root = withSlide(root, sIdx, (s) => ({
        ...s,
        shapes: replaceShape(s.shapes, target.path, withPosition(target.shape, next.x, next.y)),
      }));
    }

    const evolved = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, {
        kind: "node-updated",
        nodeId: targets[0].shape.id,
        path: ["slides", sIdx, "shapes"],
        field: "alignment",
        summary: `align ${targets.length} shapes ${payload.mode}`,
      }),
    };
  },
};

/**
 * Distribute three-or-more shapes evenly along an axis. The leftmost
 * (or topmost) and rightmost (or bottommost) shapes stay put; the
 * intermediate ones get spaced so their *centres* are equidistant.
 * Centre-based distribution matches PowerPoint's behaviour and avoids
 * weird gaps when shapes have different sizes.
 */
export const distributeShapesHandler: CommandHandler<DistributeShapesPayload, PptxSnapshot> = {
  type: "pptx:distribute-shapes",
  apply(snapshot, payload) {
    if (!payload.shapeIds || payload.shapeIds.length < 3) {
      throw makeError("invalid-payload", "distribute-shapes requires at least 3 shape ids");
    }
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const targets = collectAlignTargets(slide, payload.shapeIds);
    if (targets.length < 3) {
      throw makeError("invalid-payload", "distribute-shapes needs at least 3 distributable shapes");
    }
    const isHorizontal = payload.axis === "horizontal";
    const sorted = [...targets].sort((a, b) =>
      isHorizontal
        ? a.box.x + a.box.cx / 2 - (b.box.x + b.box.cx / 2)
        : a.box.y + a.box.cy / 2 - (b.box.y + b.box.cy / 2)
    );
    const first = sorted[0];
    const last = sorted[sorted.length - 1];
    const firstCenter = isHorizontal ? first.box.x + first.box.cx / 2 : first.box.y + first.box.cy / 2;
    const lastCenter = isHorizontal ? last.box.x + last.box.cx / 2 : last.box.y + last.box.cy / 2;
    const span = lastCenter - firstCenter;
    const step = span / (sorted.length - 1);

    let root = snapshot.root;
    for (let i = 1; i < sorted.length - 1; i++) {
      const target = sorted[i];
      const desiredCenter = firstCenter + step * i;
      const nextX = isHorizontal ? Math.round(desiredCenter - target.box.cx / 2) : target.box.x;
      const nextY = isHorizontal ? target.box.y : Math.round(desiredCenter - target.box.cy / 2);
      if (nextX === target.box.x && nextY === target.box.y) continue;
      root = withSlide(root, sIdx, (s) => ({
        ...s,
        shapes: replaceShape(s.shapes, target.path, withPosition(target.shape, nextX, nextY)),
      }));
    }

    const evolved = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, {
        kind: "node-updated",
        nodeId: sorted[0].shape.id,
        path: ["slides", sIdx, "shapes"],
        field: "distribution",
        summary: `distribute ${sorted.length} shapes ${payload.axis}`,
      }),
    };
  },
};

interface AlignTarget {
  readonly shape: Shape;
  readonly path: ReadonlyArray<number>;
  readonly box: BoundingBox;
}

function collectAlignTargets(slide: Slide, ids: ReadonlyArray<NodeId>): AlignTarget[] {
  const out: AlignTarget[] = [];
  for (const id of ids) {
    const found = findShapeInSlide(slide, id);
    if (found.shape.kind === "opaque") continue;
    const box = shapeBoundingBox(found.shape);
    if (!box) continue;
    out.push({ shape: found.shape, path: found.path, box });
  }
  return out;
}

function unionBox(boxes: ReadonlyArray<BoundingBox>): BoundingBox {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const b of boxes) {
    if (b.x < minX) minX = b.x;
    if (b.y < minY) minY = b.y;
    if (b.x + b.cx > maxX) maxX = b.x + b.cx;
    if (b.y + b.cy > maxY) maxY = b.y + b.cy;
  }
  return { x: minX, y: minY, cx: maxX - minX, cy: maxY - minY };
}

function alignBox(
  box: BoundingBox,
  union: BoundingBox,
  mode: AlignShapesPayload["mode"]
): { x: number; y: number } {
  switch (mode) {
    case "left":
      return { x: union.x, y: box.y };
    case "right":
      return { x: union.x + union.cx - box.cx, y: box.y };
    case "center-h":
      return { x: Math.round(union.x + union.cx / 2 - box.cx / 2), y: box.y };
    case "top":
      return { x: box.x, y: union.y };
    case "bottom":
      return { x: box.x, y: union.y + union.cy - box.cy };
    case "middle-v":
      return { x: box.x, y: Math.round(union.y + union.cy / 2 - box.cy / 2) };
  }
}

/**
 * Returns the shape with a refreshed `position`. The model treats
 * `position` as readonly, so we spread+overwrite. Shapes without a
 * pre-existing position (placeholders) are filtered out before we get
 * here, so the round-tripped XML always carries a well-formed `<a:off>`.
 */
function withPosition(shape: Shape, xEmu: number, yEmu: number): Shape {
  return { ...shape, position: { xEmu: Math.round(xEmu), yEmu: Math.round(yEmu) } };
}
