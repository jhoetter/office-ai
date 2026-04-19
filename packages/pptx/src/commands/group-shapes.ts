import type { CommandHandler, NodeId } from "@officeai/core";
import type { GroupShape, OpaqueXml, PptxSnapshot, Shape } from "../model/types.js";
import { buildDiff, evolveSnapshot, findSlide, makeError, maxCNvPrId, withSlide } from "./helpers.js";
import type { GroupShapesPayload, UngroupShapePayload } from "./payloads.js";

/**
 * Group two-or-more top-level shapes on a slide into a fresh
 * `GroupShape`. The group's bounding box is the union of its children's
 * positions/sizes; we synthesise its `chOff`/`chExt` to match exactly so
 * children keep their original absolute coordinates. The PowerPoint
 * round-trip happily re-emits this as a normal `<p:grpSp>`.
 */
export const groupShapesHandler: CommandHandler<GroupShapesPayload, PptxSnapshot> = {
  type: "pptx:group-shapes",
  apply(snapshot, payload, ctx) {
    if (!payload.shapeIds || payload.shapeIds.length < 2) {
      throw makeError("invalid-payload", "group-shapes requires at least 2 shape ids");
    }
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const indices = payload.shapeIds.map((id) => indexOfTopShape(slide.shapes, id));
    if (indices.some((i) => i < 0)) {
      throw makeError("not-applicable", "group-shapes only works on top-level shapes");
    }
    const targets = indices.map((i) => slide.shapes[i]!);
    for (const t of targets) {
      if (!t.position || !t.size) {
        throw makeError("not-applicable", `cannot group shape "${t.name}" without explicit position/size`);
      }
    }

    // Sort indices ascending so we keep visual order; the group will be
    // inserted at the position of the first selected shape.
    const sortedIdx = [...indices].sort((a, b) => a - b);
    const insertAt = sortedIdx[0];
    const orderedTargets = sortedIdx.map((i) => slide.shapes[i]!);

    const minX = Math.min(...orderedTargets.map((s) => s.position!.xEmu));
    const minY = Math.min(...orderedTargets.map((s) => s.position!.yEmu));
    const maxX = Math.max(...orderedTargets.map((s) => s.position!.xEmu + s.size!.cxEmu));
    const maxY = Math.max(...orderedTargets.map((s) => s.position!.yEmu + s.size!.cyEmu));
    const cx = maxX - minX;
    const cy = maxY - minY;

    const cNvPrId = maxCNvPrId(slide.shapes) + 1;
    const group: GroupShape = {
      kind: "group",
      id: ctx.mintNodeId(),
      cNvPrId,
      name: payload.name ?? `Group ${cNvPrId}`,
      position: { xEmu: minX, yEmu: minY },
      size: { cxEmu: cx, cyEmu: cy },
      chOffExtRaw: synthesiseChOffExt(minX, minY, cx, cy),
      grpSpPrTail: [],
      nvGrpSpPrTail: [],
      children: orderedTargets,
    };

    const targetIdSet = new Set(orderedTargets.map((s) => s.id));
    const remaining = slide.shapes.filter((s) => !targetIdSet.has(s.id));
    const next = remaining.slice();
    next.splice(insertAt, 0, group);

    const root = withSlide(snapshot.root, sIdx, (s) => ({ ...s, shapes: next }));
    const evolved = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, {
        kind: "node-inserted",
        nodeId: group.id,
        path: ["slides", sIdx, "shapes", insertAt],
        summary: `group ${orderedTargets.length} shapes`,
      }),
    };
  },
};

/**
 * Reverse `pptx:group-shapes`: dissolve a top-level group, re-inserting
 * its children at the group's index in the slide. Children keep their
 * absolute coordinates because group authoring made `chOff` mirror the
 * group's position (so absolute = relative).
 */
export const ungroupShapeHandler: CommandHandler<UngroupShapePayload, PptxSnapshot> = {
  type: "pptx:ungroup-shape",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const idx = indexOfTopShape(slide.shapes, payload.shapeId);
    if (idx < 0) throw makeError("not-applicable", "ungroup-shape only works on top-level groups");
    const group = slide.shapes[idx];
    if (group.kind !== "group") {
      throw makeError("not-applicable", "ungroup-shape target is not a group");
    }
    const next = slide.shapes.slice();
    next.splice(idx, 1, ...group.children);

    const root = withSlide(snapshot.root, sIdx, (s) => ({ ...s, shapes: next }));
    const evolved = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, {
        kind: "node-deleted",
        nodeId: group.id,
        path: ["slides", sIdx, "shapes", idx],
        summary: `ungroup ${group.children.length} shapes`,
      }),
    };
  },
};

function indexOfTopShape(shapes: ReadonlyArray<Shape>, id: NodeId): number {
  for (let i = 0; i < shapes.length; i++) {
    if (shapes[i].id === id) return i;
  }
  return -1;
}

/**
 * Build `<a:chOff>` and `<a:chExt>` opaque XML so the serializer emits
 * a coordinate frame identical to the group's bounding box. Children
 * keep their absolute slide coordinates after grouping/ungrouping.
 */
function synthesiseChOffExt(x: number, y: number, cx: number, cy: number): OpaqueXml[] {
  const off: OpaqueXml = {
    tag: "a:chOff",
    attrs: { x: String(x), y: String(y) },
    rawAttrs: { "@_x": String(x), "@_y": String(y) },
    subtree: [],
  };
  const ext: OpaqueXml = {
    tag: "a:chExt",
    attrs: { cx: String(cx), cy: String(cy) },
    rawAttrs: { "@_cx": String(cx), "@_cy": String(cy) },
    subtree: [],
  };
  return [off, ext];
}
