import type { CommandHandler, NodeId } from "@officeai/core";
import type {
  EntranceAnimation,
  PptxSnapshot,
  Slide,
  SlideTransition,
  TransitionKind,
} from "../model/types.js";
import { buildDiff, evolveSnapshot, findShapeInSlide, findSlide, makeError, withSlide } from "./helpers.js";
import type {
  AddShapeAnimationPayload,
  RemoveShapeAnimationPayload,
  ReorderShapeAnimationsPayload,
  SetSlideTransitionPayload,
} from "./payloads.js";

// ─── pptx:set-slide-transition ───────────────────────────────────────────

export const setSlideTransitionHandler: CommandHandler<SetSlideTransitionPayload, PptxSnapshot> = {
  type: "pptx:set-slide-transition",
  apply(snapshot, payload, ctx) {
    const { slide } = findSlide(snapshot, payload.slideIndex);
    const before = describeTransition(slide.transition);

    let nextSlide: Slide;
    if (payload.kind === "none") {
      if (!slide.transition) {
        throw makeError("no-op", "slide has no transition to remove");
      }
      // Drop the transition field entirely (and its raw blob), so the
      // serializer omits <p:transition> on the rebuilt slide.
      nextSlide = omitTransition(slide);
    } else {
      const nextKind: TransitionKind = payload.kind;
      const sameKind = slide.transition?.kind === nextKind;
      const sameSpeed = (slide.transition?.speed ?? undefined) === payload.speed;
      if (sameKind && sameSpeed) {
        throw makeError("no-op", `transition is already ${nextKind}`);
      }
      const next: SlideTransition = {
        id: slide.transition?.id ?? ctx.mintNodeId(),
        kind: nextKind,
        ...(payload.speed ? { speed: payload.speed } : {}),
        // Drop any preserved raw blob: serializer rebuilds from typed fields.
      };
      nextSlide = { ...slide, transition: next };
    }

    const root = withSlide(snapshot.root, payload.slideIndex, () => nextSlide);
    const next = evolveSnapshot(snapshot, root, {
      slides: [snapshot.root.slides[payload.slideIndex].partPath],
    });
    const after = describeTransition(nextSlide.transition);
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: nextSlide.id,
        path: ["slides", payload.slideIndex, "transition"],
        field: "kind",
        summary: `${before} → ${after}`,
      }),
    };
  },
};

// ─── pptx:add-shape-animation ────────────────────────────────────────────

export const addShapeAnimationHandler: CommandHandler<AddShapeAnimationPayload, PptxSnapshot> = {
  type: "pptx:add-shape-animation",
  apply(snapshot, payload, ctx) {
    const { slide } = findSlide(snapshot, payload.slideIndex);
    const { shape } = findShapeInSlide(slide, payload.shapeId);
    if (!shape.cNvPrId || shape.cNvPrId <= 0) {
      throw makeError("not-applicable", `shape ${payload.shapeId} has no cNvPrId; cannot be animated`);
    }
    const insertAt = clampInsertIndex(payload.at, slide.animations.length);
    const newAnim: EntranceAnimation = {
      id: ctx.mintNodeId(),
      targetCNvPrId: shape.cNvPrId,
      effect: payload.effect,
      ...(payload.durationMs !== undefined ? { durationMs: payload.durationMs } : {}),
      order: insertAt,
    };
    const newAnimations = [
      ...slide.animations.slice(0, insertAt),
      newAnim,
      ...slide.animations.slice(insertAt),
    ].map((a, i) => ({ ...a, order: i }));
    // Drop the verbatim <p:timing> tail: animations are now model-driven.
    const nextSlide = dropTimingTail({ ...slide, animations: newAnimations });

    const root = withSlide(snapshot.root, payload.slideIndex, () => nextSlide);
    const next = evolveSnapshot(snapshot, root, {
      slides: [snapshot.root.slides[payload.slideIndex].partPath],
    });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: newAnim.id,
        path: ["slides", payload.slideIndex, "animations", insertAt],
        summary: `entrance ${payload.effect} on ${payload.shapeId}`,
      }),
    };
  },
};

// ─── pptx:remove-shape-animation ─────────────────────────────────────────

export const removeShapeAnimationHandler: CommandHandler<RemoveShapeAnimationPayload, PptxSnapshot> = {
  type: "pptx:remove-shape-animation",
  apply(snapshot, payload) {
    const { slide } = findSlide(snapshot, payload.slideIndex);
    const idx = slide.animations.findIndex((a) => a.id === payload.animationId);
    if (idx < 0) {
      throw makeError(
        "unknown-target",
        `animation ${payload.animationId} not found on slide ${payload.slideIndex}`
      );
    }
    const removed = slide.animations[idx]!;
    const newAnimations = [...slide.animations.slice(0, idx), ...slide.animations.slice(idx + 1)].map(
      (a, i) => ({ ...a, order: i })
    );
    const nextSlide = dropTimingTail({ ...slide, animations: newAnimations });

    const root = withSlide(snapshot.root, payload.slideIndex, () => nextSlide);
    const next = evolveSnapshot(snapshot, root, {
      slides: [snapshot.root.slides[payload.slideIndex].partPath],
    });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-deleted",
        nodeId: removed.id,
        path: ["slides", payload.slideIndex, "animations", idx],
        summary: `entrance ${removed.effect}`,
      }),
    };
  },
};

// ─── pptx:reorder-shape-animations ───────────────────────────────────────

export const reorderShapeAnimationsHandler: CommandHandler<ReorderShapeAnimationsPayload, PptxSnapshot> = {
  type: "pptx:reorder-shape-animations",
  apply(snapshot, payload) {
    const { slide } = findSlide(snapshot, payload.slideIndex);
    if (payload.order.length !== slide.animations.length) {
      throw makeError(
        "invalid-payload",
        `order length (${payload.order.length}) must equal animations length (${slide.animations.length})`
      );
    }
    const byId = new Map<NodeId, EntranceAnimation>();
    for (const a of slide.animations) byId.set(a.id, a);
    const newAnimations: EntranceAnimation[] = [];
    const seen = new Set<NodeId>();
    for (let i = 0; i < payload.order.length; i++) {
      const id = payload.order[i]!;
      if (seen.has(id)) {
        throw makeError("invalid-payload", `duplicate id in order: ${id}`);
      }
      seen.add(id);
      const a = byId.get(id);
      if (!a) {
        throw makeError("invalid-payload", `id ${id} is not in current animations`);
      }
      newAnimations.push({ ...a, order: i });
    }
    const nextSlide = dropTimingTail({ ...slide, animations: newAnimations });

    const root = withSlide(snapshot.root, payload.slideIndex, () => nextSlide);
    const next = evolveSnapshot(snapshot, root, {
      slides: [snapshot.root.slides[payload.slideIndex].partPath],
    });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: nextSlide.id,
        path: ["slides", payload.slideIndex, "animations"],
        field: "order",
        summary: `permutation of ${payload.order.length}`,
      }),
    };
  },
};

// ─── helpers ─────────────────────────────────────────────────────────────

function describeTransition(t: SlideTransition | undefined): string {
  if (!t) return "(none)";
  return t.speed ? `${t.kind}/${t.speed}` : t.kind;
}

function omitTransition(slide: Slide): Slide {
  const { transition: _t, ...rest } = slide;
  void _t;
  return rest as Slide;
}

/**
 * After typed-animation edits, the original `<p:timing>` blob is no longer
 * authoritative. Drop it so the serializer rebuilds from `Slide.animations`.
 */
function dropTimingTail(slide: Slide): Slide {
  if (!slide.timingTailRaw) return slide;
  const { timingTailRaw: _t, ...rest } = slide;
  void _t;
  return rest as Slide;
}

function clampInsertIndex(at: number | undefined, len: number): number {
  if (at === undefined) return len;
  if (at < 0) return 0;
  if (at > len) return len;
  return at;
}
