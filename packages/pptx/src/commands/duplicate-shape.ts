import type { CommandHandler, NodeId } from "@officeai/core";
import type { PptxSnapshot, Shape, TextParagraph, TextRun } from "../model/types.js";
import {
  buildDiff,
  evolveSnapshot,
  findShapeInSlide,
  findSlide,
  makeError,
  maxCNvPrId,
  withSlide,
} from "./helpers.js";
import type { DuplicateShapePayload } from "./payloads.js";

const DEFAULT_OFFSET_EMU = 228_600; // ¼ inch — matches PowerPoint's Cmd+D nudge.

/**
 * Clone a single shape on the same slide. The clone is appended to the
 * slide's top-level shape array (so it sits in front), gets a fresh
 * `cNvPrId` so connectors / animations targeting the original aren't
 * accidentally re-pointed, and is offset by `(dxEmu, dyEmu)` (default ¼")
 * so the user can immediately see and grab it.
 *
 * Source shape:
 *   • Top-level shapes only — duplicating a child of a `GroupShape` is
 *     not supported yet (the OOXML round-trip would have to invent
 *     `<a:chOff>`-relative coordinates, which we don't yet model).
 *   • Connectors are duplicated verbatim; their endpoint anchors keep
 *     pointing to the same source/target shapes (matching PowerPoint).
 *   • `Opaque` shapes are refused — we can't safely deep-copy XML we
 *     don't introspect.
 */
export const duplicateShapeHandler: CommandHandler<DuplicateShapePayload, PptxSnapshot> = {
  type: "pptx:duplicate-shape",
  apply(snapshot, payload, ctx) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (path.length !== 1) {
      throw makeError("not-applicable", "duplicating nested group children is not supported");
    }
    if (shape.kind === "opaque") {
      throw makeError("not-applicable", "cannot duplicate opaque shapes");
    }

    const dx = Number.isFinite(payload.dxEmu) ? Math.round(payload.dxEmu!) : DEFAULT_OFFSET_EMU;
    const dy = Number.isFinite(payload.dyEmu) ? Math.round(payload.dyEmu!) : DEFAULT_OFFSET_EMU;

    const newCNvPrId = maxCNvPrId(slide.shapes) + 1;
    const clone = cloneShape(shape, dx, dy, newCNvPrId, ctx);

    const root = withSlide(snapshot.root, sIdx, (s) => ({
      ...s,
      shapes: [...s.shapes, clone],
    }));
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: clone.id,
        path: ["slides", sIdx, "shapes", slide.shapes.length],
        summary: `duplicate ${shape.kind}`,
      }),
    };
  },
};

interface MintCtx {
  readonly mintNodeId: () => NodeId;
}

function cloneShape(shape: Shape, dx: number, dy: number, newCNvPrId: number, ctx: MintCtx): Shape {
  const next: Shape = { ...shape, id: ctx.mintNodeId(), cNvPrId: newCNvPrId };
  const offset = nudgePosition(shape, dx, dy);
  // re-mint nested ids so the new sub-tree is internally unique
  switch (next.kind) {
    case "text":
      return { ...next, ...offset, txBody: cloneTextBody(next.txBody, ctx) } as Shape;
    case "pic":
      return { ...next, ...offset } as Shape;
    case "table":
      return { ...next, ...offset } as Shape;
    case "chart":
      return { ...next, ...offset } as Shape;
    case "connector": {
      // Connectors derive their bounding box from endpoints, so we
      // shift each endpoint by the offset rather than the bounding
      // box. Anchored endpoints stay anchored — they keep referencing
      // the same source/target shapes.
      const start =
        next.start.kind === "free"
          ? { ...next.start, xEmu: next.start.xEmu + dx, yEmu: next.start.yEmu + dy }
          : next.start;
      const end =
        next.end.kind === "free"
          ? { ...next.end, xEmu: next.end.xEmu + dx, yEmu: next.end.yEmu + dy }
          : next.end;
      return { ...next, ...offset, start, end } as Shape;
    }
    case "group":
      // Re-mint group children ids so the duplicated subtree doesn't
      // share NodeIds with the source group. We don't mint new
      // cNvPrIds for nested children — they live in the group's
      // private id space and don't collide with slide-level shapes.
      return {
        ...next,
        ...offset,
        children: next.children.map((c) => cloneShape(c, 0, 0, c.cNvPrId, ctx)),
      } as Shape;
    case "opaque":
      // Refused above; reachable only if the union grows.
      return { ...next, ...offset } as Shape;
  }
}

function nudgePosition(shape: Shape, dx: number, dy: number): Pick<Shape, "position"> {
  if (!shape.position) return {};
  return { position: { xEmu: shape.position.xEmu + dx, yEmu: shape.position.yEmu + dy } };
}

function cloneTextBody(
  body: import("../model/types.js").TextBody,
  ctx: MintCtx
): import("../model/types.js").TextBody {
  const paragraphs: TextParagraph[] = body.paragraphs.map((p) => ({
    ...p,
    id: ctx.mintNodeId(),
    runs: p.runs.map(
      (r): TextRun => ({
        ...r,
        id: ctx.mintNodeId(),
      })
    ),
  }));
  return { ...body, paragraphs };
}
