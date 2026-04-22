/**
 * F4 Phase 4 — Live animation playback engine.
 *
 * Drives PowerPoint-style animation timelines against an SVG slide
 * rendered by `slideToSvgString` / `<SlideCanvas>`. The renderer wraps
 * every shape in `<g class="anim-target" data-cnvprid="N">` so this
 * engine can locate targets by `ShapeAnimation.targetCNvPrId`.
 *
 * Trigger semantics (matches PowerPoint):
 *   - `onClick`        → starts a new click group; advances on user input
 *   - `withPrevious`   → starts simultaneously with the previous step
 *   - `afterPrevious`  → starts when the previous step finishes
 *
 * The engine groups animations into "click groups", each rooted at an
 * `onClick` (or the very first animation). `clickAdvance()` plays one
 * group at a time; consumers wire it to Space/Arrow/click in present
 * mode or to a "Preview" button in the editor.
 *
 * Each preset compiles to one or more Web Animations API keyframe
 * effects on the wrapping `<g>`. Transforms (translate/scale/rotate)
 * are stacked into a single `transform` declaration since SVG/CSS
 * transforms don't compose across multiple Animation objects on the
 * same element. Visibility, opacity, filter and clip-path round out
 * the supported palette.
 *
 * The implementation is intentionally pragmatic: presets that need
 * effects we can't reproduce in CSS (e.g. `colorPulse` on a flat
 * `<rect>`) fall back to a brief opacity flash so the user still sees
 * "something happened" rather than a silent gap.
 */

import type { ShapeAnimation, Slide, SlideSize } from "../model/types.js";
import { shapeBoundingBox } from "../renderer/layout/shape.js";
import { EMU_PER_PX_AT_96DPI } from "../renderer/layout/units.js";

export interface PlaybackOptions {
  /**
   * Slide dimensions in EMU, used to translate direction-relative
   * presets (flyIn, floatOut, …) into pixel offsets that match the
   * SVG viewBox the renderer produced.
   */
  readonly slideSize: SlideSize;
  /**
   * Optional callback fired whenever the engine advances to a new
   * click group, receives the 0-based index of the group that just
   * started. Useful for the editor to highlight the currently-
   * playing badge in the right rail.
   */
  readonly onGroupStart?: (groupIndex: number) => void;
  /**
   * When true, plays the entire timeline back-to-back without waiting
   * for `clickAdvance()` calls. Used by the editor's "Preview" button.
   */
  readonly autoplay?: boolean;
  /**
   * Multiplier applied to every animation's duration. `2` makes things
   * play at half speed (handy for debugging); `0.5` doubles speed.
   * Defaults to `1`.
   */
  readonly speed?: number;
}

export interface PlaybackController {
  /**
   * Apply the initial pre-animation state to the SVG: shapes that
   * carry an entrance animation start hidden. Idempotent — calling
   * it twice has the same effect as calling it once.
   */
  prepare(): void;
  /**
   * Restore every target to its baseline (visible, no transform).
   * Cancels any in-flight animations and rewinds the click cursor
   * back to group 0.
   */
  reset(): void;
  /**
   * Plays the next click group. Resolves once every animation in
   * the group (including chained `afterPrevious` tails) finishes.
   * No-op when there are no more groups.
   */
  clickAdvance(): Promise<void>;
  /** True when at least one click group has not yet been played. */
  hasMore(): boolean;
  /** Plays every remaining click group sequentially. */
  playAll(): Promise<void>;
  /**
   * Cancel any pending animations, drop event listeners and forget
   * about the SVG. Safe to call multiple times.
   */
  destroy(): void;
}

export function createPlayback(root: Element, slide: Slide, opts: PlaybackOptions): PlaybackController {
  const speed = opts.speed && opts.speed > 0 ? opts.speed : 1;
  const slideWPx = opts.slideSize.cxEmu / EMU_PER_PX_AT_96DPI;
  const slideHPx = opts.slideSize.cyEmu / EMU_PER_PX_AT_96DPI;

  // Sort by `order` so authored sequence drives playback regardless of
  // the storage order in the parsed slide model.
  const ordered = [...slide.animations].sort((a, b) => a.order - b.order);
  const groups = groupByClick(ordered);

  // Build the cNvPrId → bounding box map up front so per-frame work is
  // limited to keyframe arithmetic. We walk the shape tree (including
  // nested groups) once; missing entries fall back to slide-sized
  // defaults if the shape can't be resolved.
  const boxesByCNvPrId = collectBoxesByCNvPrId(slide);

  let groupCursor = 0;
  let activeAnimations: Animation[] = [];
  const preparedTargets = new Set<Element>();
  let destroyed = false;

  const findTarget = (cNvPrId: number): SVGGraphicsElement | null => {
    if (destroyed) return null;
    const sel = `[data-cnvprid="${cNvPrId}"]`;
    const node = root.querySelector(sel);
    return (node as SVGGraphicsElement | null) ?? null;
  };

  const prepare = (): void => {
    if (destroyed) return;
    for (const a of ordered) {
      if (a.category !== "entrance") continue;
      const el = findTarget(a.targetCNvPrId);
      if (!el) continue;
      // Shapes with multiple entrances should still hide once.
      if (preparedTargets.has(el)) continue;
      preparedTargets.add(el);
      (el as unknown as ElementCSSInlineStyle).style.visibility = "hidden";
      (el as unknown as ElementCSSInlineStyle).style.opacity = "0";
    }
  };

  const reset = (): void => {
    cancelActive();
    for (const el of preparedTargets) {
      const style = (el as unknown as ElementCSSInlineStyle).style;
      style.visibility = "";
      style.opacity = "";
      style.transform = "";
      style.transformOrigin = "";
      style.filter = "";
      style.clipPath = "";
    }
    preparedTargets.clear();
    groupCursor = 0;
  };

  const cancelActive = (): void => {
    for (const a of activeAnimations) {
      try {
        a.cancel();
      } catch {
        // The animation may already have finished or been cancelled
        // by the browser; either way there's nothing to clean up.
      }
    }
    activeAnimations = [];
  };

  const playGroup = async (groupIndex: number): Promise<void> => {
    const group = groups[groupIndex];
    if (!group) return;
    opts.onGroupStart?.(groupIndex);
    cancelActive();

    type Plan = { anim: ShapeAnimation; offset: number; duration: number };
    const plan: Plan[] = [];
    let prevStart = 0;
    let prevEnd = 0;
    for (let i = 0; i < group.length; i++) {
      const anim = group[i]!;
      let offset: number;
      if (i === 0) {
        offset = 0;
      } else if (anim.trigger === "withPrevious") {
        offset = prevStart;
      } else {
        // `afterPrevious` (and any stray `onClick` mid-group, which
        // shouldn't happen because `groupByClick` splits there) plays
        // after the previous animation completes.
        offset = prevEnd;
      }
      offset += (anim.delayMs ?? 0) * speed;
      const duration = effectiveDurationMs(anim) * speed;
      plan.push({ anim, offset, duration });
      prevStart = offset;
      prevEnd = offset + duration;
    }

    const promises: Promise<void>[] = [];
    for (const step of plan) {
      promises.push(scheduleStep(step.anim, step.offset, step.duration));
    }
    await Promise.all(promises);
  };

  const scheduleStep = async (anim: ShapeAnimation, offsetMs: number, durationMs: number): Promise<void> => {
    if (offsetMs > 0) {
      await delay(offsetMs);
    }
    if (destroyed) return;
    const el = findTarget(anim.targetCNvPrId);
    if (!el) return;
    await runEffect(el, anim, durationMs, boxesByCNvPrId, slideWPx, slideHPx, activeAnimations);
  };

  const clickAdvance = async (): Promise<void> => {
    if (destroyed || groupCursor >= groups.length) return;
    const idx = groupCursor++;
    await playGroup(idx);
  };

  const playAll = async (): Promise<void> => {
    while (hasMore()) {
      await clickAdvance();
    }
  };

  const hasMore = (): boolean => groupCursor < groups.length;

  const destroy = (): void => {
    if (destroyed) return;
    destroyed = true;
    cancelActive();
    preparedTargets.clear();
  };

  return { prepare, reset, clickAdvance, hasMore, playAll, destroy };
}

// ─── Grouping ───────────────────────────────────────────────────────────

function groupByClick(animations: ReadonlyArray<ShapeAnimation>): ShapeAnimation[][] {
  const groups: ShapeAnimation[][] = [];
  let current: ShapeAnimation[] = [];
  for (const a of animations) {
    const startsGroup = current.length === 0 || a.trigger === "onClick";
    if (startsGroup && current.length > 0) {
      groups.push(current);
      current = [];
    }
    current.push(a);
  }
  if (current.length > 0) groups.push(current);
  return groups;
}

// ─── Effect dispatch ────────────────────────────────────────────────────

interface BoxPx {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

async function runEffect(
  el: SVGGraphicsElement,
  anim: ShapeAnimation,
  durationMs: number,
  boxes: Map<number, BoxPx>,
  slideWPx: number,
  slideHPx: number,
  active: Animation[]
): Promise<void> {
  const box = boxes.get(anim.targetCNvPrId);
  // Set the transform-origin so scale/rotate pivot around the shape's
  // own centre rather than the SVG origin. CSS transform-origin on
  // SVG `<g>` honours user units (1px = 1 user unit in our viewBox).
  if (box) {
    const cx = box.x + box.cx / 2;
    const cy = box.y + box.cy / 2;
    (el as unknown as ElementCSSInlineStyle).style.transformOrigin = `${cx}px ${cy}px`;
    (el as unknown as ElementCSSInlineStyle).style.transformBox = "view-box" as string;
  }

  const { keyframes, options, postState } = buildEffect(anim, durationMs, box, slideWPx, slideHPx);
  if (!keyframes || keyframes.length === 0) return;
  const animation = el.animate(keyframes as Keyframe[], options);
  active.push(animation);
  try {
    await animation.finished;
  } catch {
    // `cancel()` rejects the finished promise; treat as a no-op.
    return;
  }
  if (postState) applyPostState(el, postState);
}

interface BuiltEffect {
  readonly keyframes: Keyframe[] | null;
  readonly options: KeyframeAnimationOptions;
  readonly postState?: PostState;
}

interface PostState {
  readonly visibility?: "visible" | "hidden";
  readonly transform?: string;
  readonly opacity?: string;
  readonly filter?: string;
  readonly clipPath?: string;
}

function applyPostState(el: SVGGraphicsElement, state: PostState): void {
  const style = (el as unknown as ElementCSSInlineStyle).style;
  if (state.visibility !== undefined) style.visibility = state.visibility;
  if (state.opacity !== undefined) style.opacity = state.opacity;
  if (state.transform !== undefined) style.transform = state.transform;
  if (state.filter !== undefined) style.filter = state.filter;
  if (state.clipPath !== undefined) style.clipPath = state.clipPath;
}

/**
 * Compile a single `ShapeAnimation` to keyframes + final state.
 *
 * The big switch is intentional: every preset that the registry can
 * round-trip must dispatch to a concrete WAAPI effect here. New
 * presets added to `presets.ts` should land a matching arm in this
 * function so the live preview stays in lockstep with what gets
 * written to OOXML.
 */
function buildEffect(
  anim: ShapeAnimation,
  durationMs: number,
  box: BoxPx | undefined,
  slideWPx: number,
  slideHPx: number
): BuiltEffect {
  const baseOpts: KeyframeAnimationOptions = {
    duration: Math.max(1, durationMs),
    easing: "ease-out",
    fill: "forwards",
  };

  switch (anim.category) {
    case "entrance":
      return buildEntrance(anim, baseOpts, box, slideWPx, slideHPx);
    case "emphasis":
      return buildEmphasis(anim, baseOpts);
    case "exit":
      return buildExit(anim, baseOpts, box, slideWPx, slideHPx);
    case "motionPath":
      return buildMotionPath(anim, baseOpts, box, slideWPx, slideHPx);
    default: {
      const _exhaustive: never = anim.category;
      void _exhaustive;
      return { keyframes: null, options: baseOpts };
    }
  }
}

function buildEntrance(
  anim: ShapeAnimation,
  opts: KeyframeAnimationOptions,
  _box: BoxPx | undefined,
  slideWPx: number,
  slideHPx: number
): BuiltEffect {
  const post: PostState = { visibility: "visible", opacity: "1", transform: "none" };
  switch (anim.preset) {
    case "appear":
      return {
        keyframes: [
          { visibility: "visible", opacity: 1 },
          { visibility: "visible", opacity: 1 },
        ],
        options: { ...opts, duration: 1 },
        postState: post,
      };
    case "fade":
      return {
        keyframes: [
          { visibility: "visible", opacity: 0 },
          { visibility: "visible", opacity: 1 },
        ],
        options: opts,
        postState: post,
      };
    case "flyIn": {
      const v = flyOffset(anim.direction ?? "left", slideWPx, slideHPx);
      return {
        keyframes: [
          {
            visibility: "visible",
            opacity: 1,
            transform: `translate(${v.dx}px, ${v.dy}px)`,
          },
          { visibility: "visible", opacity: 1, transform: "translate(0px, 0px)" },
        ],
        options: opts,
        postState: post,
      };
    }
    case "floatIn": {
      const dy = (anim.direction === "down" ? -1 : 1) * (slideHPx / 8);
      return {
        keyframes: [
          { visibility: "visible", opacity: 0, transform: `translate(0px, ${dy}px)` },
          { visibility: "visible", opacity: 1, transform: "translate(0px, 0px)" },
        ],
        options: opts,
        postState: post,
      };
    }
    case "split": {
      const horizontal = (anim.direction ?? "horizontal") !== "vertical";
      const startClip = horizontal ? "inset(0 50% 0 50%)" : "inset(50% 0 50% 0)";
      return {
        keyframes: [
          { visibility: "visible", opacity: 1, clipPath: startClip },
          { visibility: "visible", opacity: 1, clipPath: "inset(0 0 0 0)" },
        ],
        options: opts,
        postState: { ...post, clipPath: "none" },
      };
    }
    case "wipe": {
      const start = wipeStartClip(anim.direction ?? "up");
      return {
        keyframes: [
          { visibility: "visible", opacity: 1, clipPath: start },
          { visibility: "visible", opacity: 1, clipPath: "inset(0 0 0 0)" },
        ],
        options: opts,
        postState: { ...post, clipPath: "none" },
      };
    }
    case "shape": {
      const start = (anim.direction ?? "in") === "out" ? "circle(100% at 50% 50%)" : "circle(0% at 50% 50%)";
      const end = (anim.direction ?? "in") === "out" ? "circle(0% at 50% 50%)" : "circle(100% at 50% 50%)";
      return {
        keyframes: [
          { visibility: "visible", opacity: 1, clipPath: start },
          { visibility: "visible", opacity: 1, clipPath: end },
        ],
        options: opts,
        postState: { ...post, clipPath: "none" },
      };
    }
    case "wheel":
      // Approximate "wheel" with a sweeping conic clip-path; CSS
      // doesn't ship a built-in conic clip so we fall back to a
      // rotational fade-in which reads as a wheel reveal at typical
      // speeds.
      return {
        keyframes: [
          { visibility: "visible", opacity: 0, transform: "rotate(-90deg)" },
          { visibility: "visible", opacity: 1, transform: "rotate(0deg)" },
        ],
        options: opts,
        postState: post,
      };
    case "randomBars": {
      // Ten alternating bars revealed in roughly random order. We
      // approximate via a coarse staircase clip-path animation; the
      // exact pattern doesn't round-trip but the visual hint matches.
      const horizontal = (anim.direction ?? "horizontal") !== "vertical";
      const start = horizontal ? "inset(0 100% 0 0)" : "inset(100% 0 0 0)";
      return {
        keyframes: [
          { visibility: "visible", opacity: 1, clipPath: start },
          { visibility: "visible", opacity: 1, clipPath: "inset(0 0 0 0)" },
        ],
        options: opts,
        postState: { ...post, clipPath: "none" },
      };
    }
    case "growAndTurn":
      return {
        keyframes: [
          {
            visibility: "visible",
            opacity: 0,
            transform: "translate(0,0) scale(0) rotate(-180deg)",
          },
          {
            visibility: "visible",
            opacity: 1,
            transform: "translate(0,0) scale(1) rotate(0deg)",
          },
        ],
        options: opts,
        postState: post,
      };
    case "zoom":
      return {
        keyframes: [
          { visibility: "visible", opacity: 0, transform: "scale(0)" },
          { visibility: "visible", opacity: 1, transform: "scale(1)" },
        ],
        options: opts,
        postState: post,
      };
    case "swivel":
      return {
        keyframes: [
          { visibility: "visible", opacity: 1, transform: "rotateY(-180deg)" },
          { visibility: "visible", opacity: 1, transform: "rotateY(0deg)" },
        ],
        options: opts,
        postState: post,
      };
    case "bounce":
      // Bezier approximation of PowerPoint's bounce ease — drop from
      // above the slide, then settle.
      return {
        keyframes: [
          { visibility: "visible", opacity: 1, transform: `translate(0, ${-slideHPx}px)` },
          { visibility: "visible", opacity: 1, transform: "translate(0, 0)" },
        ],
        options: { ...opts, easing: "cubic-bezier(.34,1.56,.64,1)" },
        postState: post,
      };
    default:
      // Unknown entrance preset: avoid leaving the shape invisible
      // forever (its baseline visibility was set to hidden by
      // `prepare()`). A simple fade-in delivers an acceptable preview
      // until the preset gets a bespoke implementation.
      return {
        keyframes: [
          { visibility: "visible", opacity: 0 },
          { visibility: "visible", opacity: 1 },
        ],
        options: opts,
        postState: post,
      };
  }
}

function buildEmphasis(anim: ShapeAnimation, opts: KeyframeAnimationOptions): BuiltEffect {
  switch (anim.preset) {
    case "pulse":
      return {
        keyframes: [{ transform: "scale(1)" }, { transform: "scale(1.1)" }, { transform: "scale(1)" }],
        options: { ...opts, easing: "ease-in-out" },
        postState: { transform: "none" },
      };
    case "colorPulse":
      // No portable "colour pulse" without per-fill animation; fall
      // back to a brightness flash so the user sees something.
      return {
        keyframes: [{ filter: "brightness(1)" }, { filter: "brightness(1.4)" }, { filter: "brightness(1)" }],
        options: { ...opts, easing: "ease-in-out" },
        postState: { filter: "none" },
      };
    case "teeter":
      return {
        keyframes: [
          { transform: "rotate(0deg)" },
          { transform: "rotate(6deg)" },
          { transform: "rotate(-6deg)" },
          { transform: "rotate(6deg)" },
          { transform: "rotate(0deg)" },
        ],
        options: { ...opts, easing: "ease-in-out" },
        postState: { transform: "none" },
      };
    case "spin": {
      const sign = anim.direction === "counterclockwise" ? -1 : 1;
      return {
        keyframes: [{ transform: "rotate(0deg)" }, { transform: `rotate(${360 * sign}deg)` }],
        options: { ...opts, easing: "linear" },
        postState: { transform: "none" },
      };
    }
    case "growShrink":
      return {
        keyframes: [{ transform: "scale(1)" }, { transform: "scale(1.5)" }, { transform: "scale(1)" }],
        options: { ...opts, easing: "ease-in-out" },
        postState: { transform: "none" },
      };
    case "desaturate":
      return {
        keyframes: [{ filter: "saturate(1)" }, { filter: "saturate(0)" }],
        options: opts,
        postState: { filter: "saturate(0)" },
      };
    case "fontColor":
    case "lineColor":
      // Real fontColor / lineColor edits would walk into shape children
      // and animate `fill` / `stroke`. As a faithful-but-cheap stand-in
      // we flash a hue-rotate filter so something visibly changes.
      return {
        keyframes: [
          { filter: "hue-rotate(0deg)" },
          { filter: "hue-rotate(180deg)" },
          { filter: "hue-rotate(0deg)" },
        ],
        options: { ...opts, easing: "ease-in-out" },
        postState: { filter: "none" },
      };
    default:
      // Unknown emphasis preset: avoid silent gaps in preview by
      // emitting a brief brightness pulse so users see "something
      // happened" — same fallback rationale as `colorPulse`.
      return {
        keyframes: [{ filter: "brightness(1)" }, { filter: "brightness(1.3)" }, { filter: "brightness(1)" }],
        options: { ...opts, easing: "ease-in-out" },
        postState: { filter: "none" },
      };
  }
}

function buildExit(
  anim: ShapeAnimation,
  opts: KeyframeAnimationOptions,
  _box: BoxPx | undefined,
  slideWPx: number,
  slideHPx: number
): BuiltEffect {
  const post: PostState = { visibility: "hidden", opacity: "0" };
  switch (anim.preset) {
    case "disappear":
      return {
        keyframes: [
          { visibility: "visible", opacity: 1 },
          { visibility: "hidden", opacity: 0 },
        ],
        options: { ...opts, duration: 1 },
        postState: post,
      };
    case "fade":
      return {
        keyframes: [
          { visibility: "visible", opacity: 1 },
          { visibility: "visible", opacity: 0 },
        ],
        options: opts,
        postState: post,
      };
    case "flyOut": {
      const v = flyOffset(anim.direction ?? "right", slideWPx, slideHPx);
      return {
        keyframes: [
          { visibility: "visible", opacity: 1, transform: "translate(0, 0)" },
          {
            visibility: "visible",
            opacity: 1,
            transform: `translate(${v.dx}px, ${v.dy}px)`,
          },
        ],
        options: opts,
        postState: post,
      };
    }
    case "floatOut": {
      const dy = (anim.direction === "up" ? -1 : 1) * (slideHPx / 8);
      return {
        keyframes: [
          { visibility: "visible", opacity: 1, transform: "translate(0,0)" },
          { visibility: "visible", opacity: 0, transform: `translate(0, ${dy}px)` },
        ],
        options: opts,
        postState: post,
      };
    }
    case "split": {
      const horizontal = (anim.direction ?? "horizontal") !== "vertical";
      const endClip = horizontal ? "inset(0 50% 0 50%)" : "inset(50% 0 50% 0)";
      return {
        keyframes: [
          { visibility: "visible", opacity: 1, clipPath: "inset(0 0 0 0)" },
          { visibility: "visible", opacity: 1, clipPath: endClip },
        ],
        options: opts,
        postState: post,
      };
    }
    case "wipe": {
      const end = wipeStartClip(anim.direction ?? "down");
      return {
        keyframes: [
          { visibility: "visible", opacity: 1, clipPath: "inset(0 0 0 0)" },
          { visibility: "visible", opacity: 1, clipPath: end },
        ],
        options: opts,
        postState: post,
      };
    }
    case "shape": {
      const start = (anim.direction ?? "out") === "in" ? "circle(0% at 50% 50%)" : "circle(100% at 50% 50%)";
      const end = (anim.direction ?? "out") === "in" ? "circle(100% at 50% 50%)" : "circle(0% at 50% 50%)";
      return {
        keyframes: [
          { visibility: "visible", opacity: 1, clipPath: start },
          { visibility: "visible", opacity: 1, clipPath: end },
        ],
        options: opts,
        postState: post,
      };
    }
    case "wheel":
      return {
        keyframes: [
          { visibility: "visible", opacity: 1, transform: "rotate(0deg)" },
          { visibility: "visible", opacity: 0, transform: "rotate(90deg)" },
        ],
        options: opts,
        postState: post,
      };
    case "zoom":
      return {
        keyframes: [
          { visibility: "visible", opacity: 1, transform: "scale(1)" },
          { visibility: "visible", opacity: 0, transform: "scale(0)" },
        ],
        options: opts,
        postState: post,
      };
    default:
      // Unknown exit preset: end with the same hidden post-state as the
      // explicit cases, but animate via fade so the preview shows a
      // visible departure rather than a frame-1 disappearance.
      return {
        keyframes: [
          { visibility: "visible", opacity: 1 },
          { visibility: "visible", opacity: 0 },
        ],
        options: opts,
        postState: post,
      };
  }
}

function buildMotionPath(
  anim: ShapeAnimation,
  opts: KeyframeAnimationOptions,
  _box: BoxPx | undefined,
  slideWPx: number,
  slideHPx: number
): BuiltEffect {
  const path = anim.motionPath;
  if (!path) return { keyframes: null, options: opts };
  // Sample 60 evenly-spaced points along the OOXML path so WAAPI can
  // interpolate between them. We bail (no animation) when the path is
  // unparseable so a malformed string can never surface as a frozen
  // mid-animation state.
  const samples = sampleMotionPath(path, slideWPx, slideHPx, 60);
  if (samples.length < 2) return { keyframes: null, options: opts };
  const keyframes: Keyframe[] = samples.map((p, i) => ({
    transform: `translate(${p.x}px, ${p.y}px)`,
    offset: i / (samples.length - 1),
  }));
  return {
    keyframes,
    options: { ...opts, easing: "ease-in-out" },
    postState: { transform: "none" },
  };
}

// ─── Geometry helpers ───────────────────────────────────────────────────

function flyOffset(direction: string, slideWPx: number, slideHPx: number): { dx: number; dy: number } {
  switch (direction) {
    case "left":
      return { dx: -slideWPx, dy: 0 };
    case "right":
      return { dx: slideWPx, dy: 0 };
    case "up":
      return { dx: 0, dy: -slideHPx };
    case "down":
      return { dx: 0, dy: slideHPx };
    default:
      return { dx: -slideWPx, dy: 0 };
  }
}

function wipeStartClip(direction: string): string {
  // Reveal direction: start by clipping the side opposite the motion.
  switch (direction) {
    case "up":
      return "inset(100% 0 0 0)"; // hidden top, reveal upward
    case "down":
      return "inset(0 0 100% 0)";
    case "left":
      return "inset(0 0 0 100%)";
    case "right":
      return "inset(0 100% 0 0)";
    default:
      return "inset(100% 0 0 0)";
  }
}

/**
 * Walk the slide tree and capture every shape's bounding box keyed by
 * `cNvPrId` (in pixel-equivalent units, matching the SVG viewBox).
 * Uses `shapeBoundingBox` so positioned shapes inside groups are
 * resolved consistently with how the renderer painted them.
 */
function collectBoxesByCNvPrId(slide: Slide): Map<number, BoxPx> {
  const out = new Map<number, BoxPx>();
  const walk = (shapes: ReadonlyArray<{ cNvPrId: number }>): void => {
    for (const shape of shapes as ReadonlyArray<unknown>) {
      const s = shape as { cNvPrId: number; kind: string; shapes?: ReadonlyArray<unknown> };
      const bbox = shapeBoundingBox(s as Parameters<typeof shapeBoundingBox>[0]);
      if (bbox) {
        out.set(s.cNvPrId, {
          x: bbox.x / EMU_PER_PX_AT_96DPI,
          y: bbox.y / EMU_PER_PX_AT_96DPI,
          cx: bbox.cx / EMU_PER_PX_AT_96DPI,
          cy: bbox.cy / EMU_PER_PX_AT_96DPI,
        });
      }
      if (s.kind === "group" && s.shapes) {
        walk(s.shapes as ReadonlyArray<{ cNvPrId: number }>);
      }
    }
  };
  walk(slide.shapes as ReadonlyArray<{ cNvPrId: number }>);
  return out;
}

/**
 * Approximate sampler for OOXML motion paths. Tokenises the compact
 * `M / L / C / E` syntax (slide-relative units) and returns N evenly-
 * spaced (x, y) points anchored at (0, 0) — the wrapping `<g>` is
 * already positioned at the shape's origin so the keyframes only need
 * deltas. Returns an empty array for unparseable paths.
 */
function sampleMotionPath(
  path: string,
  slideWPx: number,
  slideHPx: number,
  n: number
): Array<{ x: number; y: number }> {
  const tokens = path.trim().split(/\s+/);
  const segments: Array<{
    kind: "M" | "L" | "C";
    pts: Array<{ x: number; y: number }>;
  }> = [];
  let i = 0;
  while (i < tokens.length) {
    const cmd = tokens[i++];
    if (cmd === "E" || cmd === "Z" || cmd === "z") continue;
    if (cmd === "M" || cmd === "L") {
      const x = Number(tokens[i++]);
      const y = Number(tokens[i++]);
      if (!Number.isFinite(x) || !Number.isFinite(y)) return [];
      segments.push({
        kind: cmd as "M" | "L",
        pts: [{ x: x * slideWPx, y: y * slideHPx }],
      });
    } else if (cmd === "C") {
      const xs: number[] = [];
      for (let k = 0; k < 6; k++) {
        const v = Number(tokens[i++]);
        if (!Number.isFinite(v)) return [];
        xs.push(v);
      }
      segments.push({
        kind: "C",
        pts: [
          { x: xs[0]! * slideWPx, y: xs[1]! * slideHPx },
          { x: xs[2]! * slideWPx, y: xs[3]! * slideHPx },
          { x: xs[4]! * slideWPx, y: xs[5]! * slideHPx },
        ],
      });
    } else {
      return [];
    }
  }

  // Walk the segments, accumulating arc length so we can step at a
  // uniform rate. For lines that's trivial; for cubics we approximate
  // by sampling the parameterisation densely (16 sub-samples per
  // curve) and summing.
  type Sample = { x: number; y: number; t: number };
  const polyline: Sample[] = [];
  let cursor = { x: 0, y: 0 };
  let totalLen = 0;
  const push = (x: number, y: number): void => {
    const last = polyline[polyline.length - 1];
    const segLen = last ? Math.hypot(x - last.x, y - last.y) : 0;
    totalLen += segLen;
    polyline.push({ x, y, t: totalLen });
  };
  push(0, 0);
  for (const seg of segments) {
    if (seg.kind === "M") {
      cursor = seg.pts[0]!;
      push(cursor.x, cursor.y);
    } else if (seg.kind === "L") {
      cursor = seg.pts[0]!;
      push(cursor.x, cursor.y);
    } else {
      const [c1, c2, p] = seg.pts as [
        { x: number; y: number },
        { x: number; y: number },
        { x: number; y: number },
      ];
      const start = cursor;
      for (let k = 1; k <= 16; k++) {
        const t = k / 16;
        const it = 1 - t;
        const x = it * it * it * start.x + 3 * it * it * t * c1.x + 3 * it * t * t * c2.x + t * t * t * p.x;
        const y = it * it * it * start.y + 3 * it * it * t * c1.y + 3 * it * t * t * c2.y + t * t * t * p.y;
        push(x, y);
      }
      cursor = p;
    }
  }
  if (totalLen === 0 || polyline.length < 2) return [];

  const out: Array<{ x: number; y: number }> = [];
  for (let k = 0; k < n; k++) {
    const target = (k / (n - 1)) * totalLen;
    let lo = 0;
    let hi = polyline.length - 1;
    while (lo + 1 < hi) {
      const mid = (lo + hi) >> 1;
      if (polyline[mid]!.t <= target) lo = mid;
      else hi = mid;
    }
    const a = polyline[lo]!;
    const b = polyline[hi]!;
    const span = b.t - a.t;
    const f = span === 0 ? 0 : (target - a.t) / span;
    out.push({ x: a.x + (b.x - a.x) * f, y: a.y + (b.y - a.y) * f });
  }
  return out;
}

function effectiveDurationMs(anim: ShapeAnimation): number {
  if (anim.durationMs && anim.durationMs > 0) return anim.durationMs;
  // Sensible fallback: the registry's default would be more accurate
  // but we don't import it here to keep this module side-effect-free
  // when a shape has been authored in another tool with no explicit
  // duration. 500 ms is PowerPoint's gallery default for entrance/exit.
  switch (anim.category) {
    case "entrance":
    case "exit":
      return 500;
    case "emphasis":
      return 1000;
    case "motionPath":
      return 2000;
    default:
      return 500;
  }
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => {
    if (ms <= 0) {
      resolve();
      return;
    }
    setTimeout(resolve, ms);
  });
}
