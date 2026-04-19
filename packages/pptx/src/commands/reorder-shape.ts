import type { CommandHandler } from "@officeai/core";
import type { PptxSnapshot, Shape, Slide } from "../model/types.js";
import { buildDiff, evolveSnapshot, findSlide, makeError, withSlide } from "./helpers.js";
import type { ReorderShapePayload, ReorderShapeMode } from "./payloads.js";

/**
 * D5 — z-order mutation.
 *
 * Operates on the slide's top-level shape array (and its enclosing
 * group, if the target lives inside one). PPTX paint order = document
 * order: index 0 is back-most, last index is front-most. We never
 * reorder children across a group boundary; the operation stays in
 * the same parent so grouping semantics are preserved.
 *
 * Modes:
 *   "to-front"   — move to the very end of the parent's shapes
 *   "to-back"    — move to the very beginning
 *   "forward"    — swap with the next shape (no-op at the front)
 *   "backward"   — swap with the previous shape (no-op at the back)
 */
export const reorderShapeHandler: CommandHandler<ReorderShapePayload, PptxSnapshot> = {
  type: "pptx:reorder-shape",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const targetPath = locateShape(slide.shapes, payload.shapeId);
    if (!targetPath) {
      throw makeError("unknown-target", `shape ${payload.shapeId} not found on slide`);
    }
    const nextRoot = withSlide(snapshot.root, sIdx, (s) => ({
      ...s,
      shapes: reorderAt(s.shapes, targetPath, payload.mode),
    }));
    const evolved = evolveSnapshot(snapshot, nextRoot, { slides: [slide.partPath] });
    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, {
        kind: "node-moved",
        nodeId: payload.shapeId,
        from: ["slides", sIdx, "shapes", ...targetPath],
        to: ["slides", sIdx, "shapes", ...targetPath],
        summary: payload.mode,
      }),
    };
  },
};

function locateShape(shapes: ReadonlyArray<Shape>, id: string): number[] | null {
  for (let i = 0; i < shapes.length; i++) {
    const s = shapes[i];
    if (s.id === id) return [i];
    if (s.kind === "group") {
      const inner = locateShape(s.children, id);
      if (inner) return [i, ...inner];
    }
  }
  return null;
}

function reorderAt(
  shapes: ReadonlyArray<Shape>,
  path: ReadonlyArray<number>,
  mode: ReorderShapeMode
): Shape[] {
  if (path.length === 0) throw new Error("empty path");
  const [head, ...tail] = path;
  if (tail.length === 0) {
    return reorderInArray([...shapes], head, mode);
  }
  const out = [...shapes];
  const parent = out[head];
  if (parent.kind !== "group") throw new Error("path expects group");
  out[head] = {
    ...parent,
    children: reorderAt(parent.children, tail, mode),
  };
  return out;
}

function reorderInArray(arr: Shape[], idx: number, mode: ReorderShapeMode): Shape[] {
  const last = arr.length - 1;
  switch (mode) {
    case "to-front": {
      if (idx === last) return arr;
      const [picked] = arr.splice(idx, 1);
      arr.push(picked);
      return arr;
    }
    case "to-back": {
      if (idx === 0) return arr;
      const [picked] = arr.splice(idx, 1);
      arr.unshift(picked);
      return arr;
    }
    case "forward": {
      if (idx === last) return arr;
      [arr[idx], arr[idx + 1]] = [arr[idx + 1], arr[idx]];
      return arr;
    }
    case "backward": {
      if (idx === 0) return arr;
      [arr[idx], arr[idx - 1]] = [arr[idx - 1], arr[idx]];
      return arr;
    }
    default: {
      const exhaust: never = mode;
      throw makeError("invalid-payload", `unknown reorder mode: ${String(exhaust)}`);
    }
  }
}

// Re-export for convenience.
export type { ReorderShapeMode } from "./payloads.js";

// Internal helper used by the slide editor (and the bus's diff
// emitter); not exported from the package barrel.
export function _internalLocateShape(
  shapes: ReadonlyArray<Shape>,
  id: string,
  _slide?: Slide
): number[] | null {
  return locateShape(shapes, id);
}
