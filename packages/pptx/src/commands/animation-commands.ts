import type { CommandHandler, NodeId } from "@officeai/core";
import type {
  AnimationCategory,
  AnimationDirection,
  AnimationPreset,
  AnimationTrigger,
  OpaqueXml,
  PptxSnapshot,
  ShapeAnimation,
  Slide,
  SlideTransition,
  TransitionKind,
  TransitionSpeed,
} from "../model/types.js";
import { buildDiff, evolveSnapshot, findShapeInSlide, findSlide, makeError, withSlide } from "./helpers.js";
import { findPreset } from "../animation/presets.js";
import type {
  AddShapeAnimationPayload,
  RemoveShapeAnimationPayload,
  ReorderShapeAnimationsPayload,
  SetShapeAnimationPayload,
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
      // Delta-merge: when the source carried an opaque raw blob,
      // patch only the kind / speed inside it. Everything else
      // (advanceClick, advanceTime, sound effects, unmodeled
      // children) survives. When there's no raw blob (typed
      // authoring from scratch), fall back to a clean rebuild.
      const prevRaw = slide.transition?.raw;
      const nextRaw = prevRaw ? mergeTransitionRaw(prevRaw, nextKind, payload.speed) : undefined;
      const next: SlideTransition = {
        id: slide.transition?.id ?? ctx.mintNodeId(),
        kind: nextKind,
        ...(payload.speed ? { speed: payload.speed } : {}),
        ...(nextRaw ? { raw: nextRaw } : {}),
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
    const resolved = resolvePresetFromPayload(payload);
    if (!resolved) {
      throw makeError(
        "invalid-payload",
        "must supply category+preset, or legacy effect, that resolves to a known preset"
      );
    }
    const spec = findPreset(resolved.category, resolved.preset);
    const insertAt = clampInsertIndex(payload.at, slide.animations.length);
    const newAnim: ShapeAnimation = {
      id: ctx.mintNodeId(),
      targetCNvPrId: shape.cNvPrId,
      category: resolved.category,
      preset: resolved.preset,
      trigger: payload.trigger ?? "onClick",
      ...(payload.direction
        ? { direction: payload.direction as AnimationDirection }
        : spec?.defaultDirection
          ? { direction: spec.defaultDirection }
          : {}),
      ...(payload.durationMs !== undefined
        ? { durationMs: payload.durationMs }
        : spec
          ? { durationMs: spec.defaultDurationMs }
          : {}),
      ...(payload.delayMs !== undefined ? { delayMs: payload.delayMs } : {}),
      ...(payload.motionPath !== undefined ? { motionPath: payload.motionPath } : {}),
      order: insertAt,
    };
    const newAnimations = [
      ...slide.animations.slice(0, insertAt),
      newAnim,
      ...slide.animations.slice(insertAt),
    ].map((a, i) => ({ ...a, order: i }));
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
        summary: `${resolved.category} ${resolved.preset} on ${payload.shapeId}`,
      }),
    };
  },
};

// ─── pptx:set-shape-animation ────────────────────────────────────────────
//
// Patches an existing animation in place. Mirrors the partial-update
// semantics used by other typed-edit commands: any field omitted is
// preserved, and `null` clears the optional field. Mutating the typed
// animation also drops the captured `raw` blob since the OOXML body
// no longer matches the typed values.

export const setShapeAnimationHandler: CommandHandler<SetShapeAnimationPayload, PptxSnapshot> = {
  type: "pptx:set-shape-animation",
  apply(snapshot, payload) {
    const { slide } = findSlide(snapshot, payload.slideIndex);
    const idx = slide.animations.findIndex((a) => a.id === payload.animationId);
    if (idx < 0) {
      throw makeError(
        "unknown-target",
        `animation ${payload.animationId} not found on slide ${payload.slideIndex}`
      );
    }
    const current = slide.animations[idx]!;
    const nextCategory: AnimationCategory = payload.category ?? current.category;
    const nextPreset: AnimationPreset = (payload.preset as AnimationPreset | undefined) ?? current.preset;
    if (payload.category && !payload.preset) {
      throw makeError("invalid-payload", "category change requires explicit preset");
    }
    const spec = findPreset(nextCategory, nextPreset);
    if (!spec) {
      throw makeError("invalid-payload", `unknown preset ${nextCategory}/${nextPreset}`);
    }
    const nextTrigger: AnimationTrigger = payload.trigger ?? current.trigger;

    const nextAnim: ShapeAnimation = {
      id: current.id,
      targetCNvPrId: current.targetCNvPrId,
      category: nextCategory,
      preset: nextPreset,
      trigger: nextTrigger,
      ...patchOptional("direction", payload.direction, current.direction),
      ...patchOptional("durationMs", payload.durationMs, current.durationMs),
      ...patchOptional("delayMs", payload.delayMs, current.delayMs),
      ...patchOptional("motionPath", payload.motionPath, current.motionPath),
      order: current.order,
      // raw is intentionally cleared on mutation — the typed fields are
      // now authoritative and the captured blob is stale.
    };
    if (
      shallowEqualAnim(current, nextAnim) &&
      // raw=undefined && current.raw=undefined is a no-op
      !current.raw
    ) {
      throw makeError("no-op", "animation already matches the requested state");
    }

    const newAnimations = [...slide.animations.slice(0, idx), nextAnim, ...slide.animations.slice(idx + 1)];
    const nextSlide = dropTimingTail({ ...slide, animations: newAnimations });

    const root = withSlide(snapshot.root, payload.slideIndex, () => nextSlide);
    const next = evolveSnapshot(snapshot, root, {
      slides: [snapshot.root.slides[payload.slideIndex].partPath],
    });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: nextAnim.id,
        path: ["slides", payload.slideIndex, "animations", idx],
        field: "preset",
        summary: `${nextCategory}/${nextPreset}`,
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
        summary: `${removed.category} ${removed.preset}`,
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
    const byId = new Map<NodeId, ShapeAnimation>();
    for (const a of slide.animations) byId.set(a.id, a);
    const newAnimations: ShapeAnimation[] = [];
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
 * After typed-animation edits, the original `<p:timing>` blob is
 * preserved as `timingTailRaw`. The serializer is responsible for
 * **surgically merging** the typed `animations[]` into the captured
 * tail tree: matching `<p:par>` carriers are replaced with
 * rebuilt versions for edited animations, kept verbatim for
 * preserved ones (`a.raw !== undefined`), and dropped for typed
 * animations that no longer exist. Anything outside the
 * typed-animation `<p:par>` set (mainSeq envelope, sound effects,
 * unmodelled emphasis / exit / motionPath siblings) survives
 * verbatim.
 *
 * This used to drop `timingTailRaw` entirely on any typed edit —
 * causing all unmodelled timing content to be silently lost. See
 * `docs/build-log/pptx.md` "Known issues" for the bug history,
 * and `serializer/serialize.ts` `mergeTimingFromAnimations` for
 * the merge implementation.
 */
function dropTimingTail(slide: Slide): Slide {
  // Preserve the captured tail so the serializer can merge typed
  // edits into it surgically. The previous behaviour stripped
  // `timingTailRaw` here; that strategy is now incorrect.
  return slide;
}

/**
 * Patch a captured `<p:transition>` opaque blob to reflect a new
 * typed kind/speed without losing other attributes (e.g.
 * `advClick`, `advTm`) or sibling children (sound effects). The
 * kind is encoded as a single child element (`<p:fade/>`,
 * `<p:push dir="r"/>`, …); we swap that child while leaving
 * everything else byte-identical.
 *
 * Returns `undefined` when the new kind is `unsupported` — the
 * caller will fall back to a clean rebuild because we cannot
 * synthesise an arbitrary kind element from the typed model.
 */
function mergeTransitionRaw(
  raw: OpaqueXml,
  nextKind: TransitionKind,
  nextSpeed: TransitionSpeed | undefined
): OpaqueXml | undefined {
  if (nextKind === "unsupported") return undefined;

  // Update attribute maps. PowerPoint omits `spd` for default
  // ("med"); we preserve absence vs presence per the typed value.
  const attrs: Record<string, string> = { ...raw.attrs };
  const rawAttrs: Record<string, string> = { ...raw.rawAttrs };
  if (nextSpeed) {
    attrs.spd = nextSpeed;
    rawAttrs["@_spd"] = nextSpeed;
  } else {
    delete attrs.spd;
    delete rawAttrs["@_spd"];
  }

  // Swap the kind element inside subtree. Recognised kind tags are
  // the same set the parser walks; any element with a tag in that
  // set is replaced by `<p:${nextKind}/>` (no attrs). All other
  // children (e.g. `<p:sndAc>` for sound effects) survive verbatim.
  const newKindTag = `p:${nextKind}`;
  let replaced = false;
  const subtree: unknown[] = [];
  for (const node of raw.subtree) {
    const tag = nodeTag(node);
    if (tag && KNOWN_TRANSITION_KIND_TAGS.has(tag)) {
      if (!replaced) {
        subtree.push({ [newKindTag]: [] });
        replaced = true;
      }
      // Drop any extra kind children — there should only be one.
      continue;
    }
    subtree.push(node);
  }
  if (!replaced) {
    // Source had no recognised kind child (rare — `unsupported`
    // case parsed earlier). Insert the new one at the head so it
    // matches PowerPoint's emission order.
    subtree.unshift({ [newKindTag]: [] });
  }

  return { tag: raw.tag, attrs, rawAttrs, subtree };
}

const KNOWN_TRANSITION_KIND_TAGS = new Set<string>([
  "p:fade",
  "p:push",
  "p:wipe",
  "p:cut",
  "p:split",
  "p:cover",
  "p:zoom",
  "p:dissolve",
  "p:circle",
  "p:diamond",
  "p:plus",
  "p:newsflash",
  "p:pull",
  "p:random",
  "p:strips",
  "p:wedge",
  "p:wheel",
  "p:checker",
  "p:blinds",
]);

function nodeTag(node: unknown): string | undefined {
  if (!node || typeof node !== "object") return undefined;
  const keys = Object.keys(node);
  for (const k of keys) {
    if (k.startsWith(":") || k.startsWith("#")) continue;
    return k;
  }
  return undefined;
}

function clampInsertIndex(at: number | undefined, len: number): number {
  if (at === undefined) return len;
  if (at < 0) return 0;
  if (at > len) return len;
  return at;
}

/**
 * Resolve the `(category, preset)` pair from an `addShapeAnimation`
 * payload. Accepts the new typed fields, the legacy `effect` shortcut,
 * or just a `preset` (defaulting to entrance). Returns `null` if the
 * payload is too ambiguous to dispatch on.
 */
function resolvePresetFromPayload(
  payload: AddShapeAnimationPayload
): { category: AnimationCategory; preset: AnimationPreset } | null {
  if (payload.category && payload.preset) {
    return { category: payload.category, preset: payload.preset as AnimationPreset };
  }
  if (payload.preset) {
    // Default to entrance when only the preset key was supplied.
    return { category: "entrance", preset: payload.preset as AnimationPreset };
  }
  if (payload.effect) {
    const map: Record<string, AnimationPreset> = {
      appear: "appear",
      fade: "fade",
      "fly-in": "flyIn",
      wipe: "wipe",
    };
    const preset = map[payload.effect];
    if (preset) return { category: "entrance", preset };
  }
  return null;
}

/**
 * Build a partial-update payload entry. `null` means "clear", `undefined`
 * means "preserve current value", any other value is an explicit update.
 */
function patchOptional<K extends string, V>(
  key: K,
  patch: V | null | undefined,
  current: V | undefined
): Partial<Record<K, V>> {
  if (patch === null) return {};
  if (patch === undefined) return current === undefined ? {} : ({ [key]: current } as Partial<Record<K, V>>);
  return { [key]: patch } as Partial<Record<K, V>>;
}

function shallowEqualAnim(a: ShapeAnimation, b: ShapeAnimation): boolean {
  return (
    a.targetCNvPrId === b.targetCNvPrId &&
    a.category === b.category &&
    a.preset === b.preset &&
    a.trigger === b.trigger &&
    a.direction === b.direction &&
    a.durationMs === b.durationMs &&
    a.delayMs === b.delayMs &&
    a.motionPath === b.motionPath
  );
}
