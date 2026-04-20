import type { CommandHandler } from "@officeai/core";
import type { Picture, PictureSrcRect, PptxSnapshot } from "../model/types.js";
import {
  buildDiff,
  evolveSnapshot,
  findShapeInSlide,
  findSlide,
  makeError,
  replaceShape,
  withSlide,
} from "./helpers.js";
import type { CropPicturePayload } from "./payloads.js";

/**
 * D7 — write the `<a:srcRect>` crop box on a `Picture` shape. The
 * payload carries percentages (0–100) for ergonomics; the
 * serializer is responsible for the 1000-multiplied OOXML units.
 *
 * Idempotent: passing `0/0/0/0` clears an existing crop (the
 * model's `srcRect` field is dropped entirely so the serializer
 * doesn't emit an empty `<a:srcRect>`).
 */
export const cropPictureHandler: CommandHandler<CropPicturePayload, PptxSnapshot> = {
  type: "pptx:crop-picture",
  apply(snapshot, payload) {
    validate(payload);

    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (shape.kind !== "pic") {
      throw makeError(
        "invalid-target",
        `shape ${payload.shapeId} is not a picture (kind=${shape.kind})`
      );
    }

    const next = nextSrcRect(payload);
    const updated: Picture = next ? { ...shape, srcRect: next } : stripSrcRect(shape);

    const root = withSlide(snapshot.root, sIdx, (s) => ({
      ...s,
      shapes: replaceShape(s.shapes, path, updated),
    }));
    const out = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next: out,
      diff: buildDiff(snapshot.revision, out.revision, {
        kind: "node-updated",
        nodeId: shape.id,
        path: ["slides", sIdx, "shapes", ...path, "srcRect"],
        field: "srcRect",
        summary: next
          ? `crop l=${next.leftPct}% t=${next.topPct}% r=${next.rightPct}% b=${next.bottomPct}%`
          : "clear-crop",
      }),
    };
  },
};

function validate(payload: CropPicturePayload): void {
  const { leftPct, topPct, rightPct, bottomPct } = payload;
  for (const [name, value] of [
    ["leftPct", leftPct],
    ["topPct", topPct],
    ["rightPct", rightPct],
    ["bottomPct", bottomPct],
  ] as const) {
    if (!Number.isFinite(value)) {
      throw makeError("invalid-payload", `${name} must be a finite number`);
    }
    if (value < 0 || value > 100) {
      throw makeError("invalid-payload", `${name} must be in [0, 100], got ${value}`);
    }
  }
  if (leftPct + rightPct >= 100) {
    throw makeError(
      "invalid-payload",
      `leftPct + rightPct must be < 100 (got ${leftPct} + ${rightPct} = ${leftPct + rightPct})`
    );
  }
  if (topPct + bottomPct >= 100) {
    throw makeError(
      "invalid-payload",
      `topPct + bottomPct must be < 100 (got ${topPct} + ${bottomPct} = ${topPct + bottomPct})`
    );
  }
}

function nextSrcRect(payload: CropPicturePayload): PictureSrcRect | null {
  if (
    payload.leftPct === 0 &&
    payload.topPct === 0 &&
    payload.rightPct === 0 &&
    payload.bottomPct === 0
  ) {
    return null;
  }
  return {
    leftPct: payload.leftPct,
    topPct: payload.topPct,
    rightPct: payload.rightPct,
    bottomPct: payload.bottomPct,
  };
}

function stripSrcRect(shape: Picture): Picture {
  if (!shape.srcRect) return shape;
  const { srcRect: _drop, ...rest } = shape;
  return rest;
}
