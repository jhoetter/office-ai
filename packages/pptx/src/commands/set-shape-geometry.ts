import type { CommandHandler } from "@officeai/core";
import type { OpaqueXml, PptxSnapshot, TextShape, Picture } from "../model/types.js";
import {
  buildDiff,
  evolveSnapshot,
  findShapeInSlide,
  findSlide,
  makeError,
  replaceShape,
  withSlide,
} from "./helpers.js";

export interface SetShapeGeometryPayload {
  readonly slideIndex: number;
  readonly shapeId: string;
  /**
   * Adjustment name (e.g. `adj`, `adj1`, `adj2`). Each preset has its
   * own set of adjustable parameters; for `roundRect` the only knob
   * is `adj` (corner radius, 0–50000 representing 0–50% of min(w,h)).
   */
  readonly adjName: string;
  /**
   * OOXML adjustment value in 1000-th-of-percent units (i.e. 0–100000
   * for percentage adjustments, or 60_000-th-of-degree for angular
   * ones). Pass `null` to clear an explicit adjustment and revert to
   * the preset default.
   */
  readonly value: number | null;
}

/**
 * Mutate a single `<a:gd>` adjustment inside a shape's
 * `<a:prstGeom>/<a:avLst>`. Used by the on-canvas yellow handle
 * (PowerPoint convention) and the Shape Format ribbon's per-preset
 * sliders.
 *
 * Round-trip discipline: only the touched `<a:gd>` entry is rewritten
 * — every other `<a:avLst>` child and every other `spPrTail` child
 * survives byte-clean.
 */
export const setShapeGeometryHandler: CommandHandler<SetShapeGeometryPayload, PptxSnapshot> = {
  type: "pptx:set-shape-geometry",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (shape.kind !== "text" && shape.kind !== "pic") {
      throw makeError("not-applicable", `cannot set geometry adjustment on shape of kind ${shape.kind}`);
    }
    const next = applyAdjustment(shape, payload.adjName, payload.value);
    if (next === shape) {
      throw makeError("no-op", `${payload.adjName} already at ${payload.value}`);
    }
    const root = withSlide(snapshot.root, sIdx, (s) => ({
      ...s,
      shapes: replaceShape(s.shapes, path, next),
    }));
    const evolved = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });
    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, {
        kind: "node-updated",
        nodeId: shape.id,
        path: ["slides", sIdx, "shapes", ...path],
        field: "geometry",
        summary: `${payload.adjName}=${payload.value ?? "(default)"}`,
      }),
    };
  },
};

type GeometryShape = TextShape | Picture;

function applyAdjustment(shape: GeometryShape, adjName: string, value: number | null): GeometryShape {
  const tail = shape.spPrTail;
  const idx = tail.findIndex((c) => c.tag === "a:prstGeom");
  if (idx < 0) {
    // No `<a:prstGeom>` to adjust. Some shapes (custGeom, opaque
    // children-only) can't be tuned with adjustments — bail without
    // throwing so the caller can fall back to a generic resize.
    return shape;
  }
  const prst = tail[idx]!;
  const avLstIdx = prst.subtree.findIndex(
    (n) =>
      n &&
      typeof n === "object" &&
      !Array.isArray(n) &&
      Object.keys(n as Record<string, unknown>).filter((k) => k !== ":@")[0] === "a:avLst"
  );
  let nextSubtree: unknown[];
  if (avLstIdx < 0) {
    if (value === null) return shape; // already implicit default
    nextSubtree = [...prst.subtree, makeAvLstWithGd(adjName, value)];
  } else {
    const avLstNode = prst.subtree[avLstIdx] as Record<string, unknown>;
    const gds = (avLstNode["a:avLst"] as unknown[] | undefined) ?? [];
    const gdIdx = gds.findIndex((g) => {
      if (!g || typeof g !== "object" || Array.isArray(g)) return false;
      const obj = g as Record<string, unknown>;
      if (Object.keys(obj).filter((k) => k !== ":@")[0] !== "a:gd") return false;
      const a = (obj[":@"] as Record<string, unknown> | undefined) ?? {};
      return a["@_name"] === adjName;
    });
    let nextGds: unknown[];
    if (value === null) {
      if (gdIdx < 0) return shape;
      nextGds = [...gds.slice(0, gdIdx), ...gds.slice(gdIdx + 1)];
    } else {
      const newGd = {
        "a:gd": [],
        ":@": { "@_name": adjName, "@_fmla": `val ${Math.round(value)}` },
      };
      nextGds = gdIdx < 0 ? [...gds, newGd] : [...gds.slice(0, gdIdx), newGd, ...gds.slice(gdIdx + 1)];
    }
    const nextAvLst = { "a:avLst": nextGds };
    nextSubtree =
      nextGds.length === 0
        ? [...prst.subtree.slice(0, avLstIdx), ...prst.subtree.slice(avLstIdx + 1)]
        : [...prst.subtree.slice(0, avLstIdx), nextAvLst, ...prst.subtree.slice(avLstIdx + 1)];
  }
  const nextPrst: OpaqueXml = { ...prst, subtree: nextSubtree };
  const nextTail = [...tail.slice(0, idx), nextPrst, ...tail.slice(idx + 1)];
  return { ...shape, spPrTail: nextTail };
}

function makeAvLstWithGd(name: string, value: number): Record<string, unknown> {
  return {
    "a:avLst": [
      {
        "a:gd": [],
        ":@": { "@_name": name, "@_fmla": `val ${Math.round(value)}` },
      },
    ],
  };
}
