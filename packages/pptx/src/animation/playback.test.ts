import { describe, expect, it } from "vitest";
import { createPlayback } from "./playback.js";
import type { ShapeAnimation, Slide, SlideSize } from "../model/types.js";

/**
 * Hand-rolled minimal DOM stub. The playback engine only touches a
 * tiny surface area:
 *   - `root.querySelector("[data-cnvprid='N']")` to look up targets
 *   - `el.style.{visibility,opacity,transform,…}` for prepare/reset
 *   - `el.animate(keyframes, options)` returning an Animation-like
 *     object with a `.finished` Promise + `addEventListener` + `cancel`
 *
 * That's it — no need to spin up jsdom. Keeping the harness in this
 * file means the playback contract (prepare hides every entrance
 * target, clickAdvance plays one group at a time, hasMore flips false
 * at the end) is locked down by a fast Node-only vitest.
 */
function makeFakeAnimation(): {
  animation: Animation;
  resolve: () => void;
  cancelled: () => boolean;
} {
  let resolveFinished: () => void;
  const finished = new Promise<void>((r) => {
    resolveFinished = r;
  }) as Promise<Animation> & Promise<void>;
  let isCancelled = false;
  const animation = {
    finished,
    addEventListener: (_type: string, listener: () => void) => {
      finished.then(() => listener());
    },
    cancel: () => {
      isCancelled = true;
      resolveFinished();
    },
  } as unknown as Animation;
  return { animation, resolve: () => resolveFinished(), cancelled: () => isCancelled };
}

interface FakeStyle {
  visibility?: string;
  opacity?: string;
  transform?: string;
  transformOrigin?: string;
  transformBox?: string;
  filter?: string;
  clipPath?: string;
}

interface FakeTarget {
  cNvPrId: number;
  style: FakeStyle;
  animations: Array<{ keyframes: unknown; options: unknown }>;
  pendingResolve: (() => void) | null;
}

function makeRoot(targets: ReadonlyArray<FakeTarget>): Element {
  const byId = new Map<number, FakeTarget>();
  for (const t of targets) byId.set(t.cNvPrId, t);
  const querySelector = (sel: string): unknown => {
    const m = sel.match(/data-cnvprid="(\d+)"/);
    if (!m) return null;
    const id = Number(m[1]);
    const target = byId.get(id);
    if (!target) return null;
    return {
      style: target.style,
      animate: (keyframes: unknown, options: unknown) => {
        target.animations.push({ keyframes, options });
        const fa = makeFakeAnimation();
        target.pendingResolve = fa.resolve;
        // Resolve on next microtask so the await actually yields —
        // matches WAAPI, where finished resolves asynchronously.
        queueMicrotask(() => fa.resolve());
        return fa.animation;
      },
    };
  };
  return { querySelector } as unknown as Element;
}

function fadeAnimation(
  targetCNvPrId: number,
  order: number,
  trigger: ShapeAnimation["trigger"]
): ShapeAnimation {
  return {
    id: `anim-${order}`,
    targetCNvPrId,
    category: "entrance",
    preset: "fade",
    presetClass: "entr",
    trigger,
    order,
    durationMs: 100,
  };
}

const SLIDE_SIZE: SlideSize = { cxEmu: 9144000, cyEmu: 6858000 };

function makeSlide(animations: ShapeAnimation[]): Slide {
  // Minimal Slide skeleton — playback only needs `animations` + a
  // `shapes` walker (collectBoxesByCNvPrId tolerates an empty list).
  return {
    id: "slide-1",
    partPath: "ppt/slides/slide1.xml",
    cSld: { name: undefined },
    shapes: [],
    animations,
    transition: undefined,
    notesSlidePartPath: undefined,
    relationships: [],
    rawSpTree: undefined,
    background: undefined,
    layoutPartPath: undefined,
    masterPartPath: undefined,
    extLst: undefined,
  } as unknown as Slide;
}

describe("playback engine", () => {
  it("prepare() hides every entrance target", () => {
    const t1: FakeTarget = { cNvPrId: 11, style: {}, animations: [], pendingResolve: null };
    const t2: FakeTarget = { cNvPrId: 22, style: {}, animations: [], pendingResolve: null };
    const slide = makeSlide([fadeAnimation(11, 1, "onClick"), fadeAnimation(22, 2, "onClick")]);
    const root = makeRoot([t1, t2]);
    const ctrl = createPlayback(root, slide, { slideSize: SLIDE_SIZE });
    ctrl.prepare();
    expect(t1.style.visibility).toBe("hidden");
    expect(t1.style.opacity).toBe("0");
    expect(t2.style.visibility).toBe("hidden");
    expect(t2.style.opacity).toBe("0");
    ctrl.destroy();
  });

  it("clickAdvance() plays one click-group at a time and hasMore flips false", async () => {
    const t1: FakeTarget = { cNvPrId: 11, style: {}, animations: [], pendingResolve: null };
    const t2: FakeTarget = { cNvPrId: 22, style: {}, animations: [], pendingResolve: null };
    const slide = makeSlide([fadeAnimation(11, 1, "onClick"), fadeAnimation(22, 2, "onClick")]);
    const root = makeRoot([t1, t2]);
    const ctrl = createPlayback(root, slide, { slideSize: SLIDE_SIZE });
    ctrl.prepare();

    expect(ctrl.hasMore()).toBe(true);
    await ctrl.clickAdvance();
    expect(t1.animations.length).toBe(1);
    expect(t2.animations.length).toBe(0);
    expect(ctrl.hasMore()).toBe(true);

    await ctrl.clickAdvance();
    expect(t2.animations.length).toBe(1);
    expect(ctrl.hasMore()).toBe(false);

    // A further advance should be a no-op, not a throw.
    await ctrl.clickAdvance();
    expect(t1.animations.length).toBe(1);
    expect(t2.animations.length).toBe(1);

    ctrl.destroy();
  });

  it("withPrevious co-fires inside one click group", async () => {
    const t1: FakeTarget = { cNvPrId: 11, style: {}, animations: [], pendingResolve: null };
    const t2: FakeTarget = { cNvPrId: 22, style: {}, animations: [], pendingResolve: null };
    const slide = makeSlide([fadeAnimation(11, 1, "onClick"), fadeAnimation(22, 2, "withPrevious")]);
    const root = makeRoot([t1, t2]);
    const ctrl = createPlayback(root, slide, { slideSize: SLIDE_SIZE });
    ctrl.prepare();

    await ctrl.clickAdvance();
    expect(t1.animations.length).toBe(1);
    expect(t2.animations.length).toBe(1);
    expect(ctrl.hasMore()).toBe(false);
    ctrl.destroy();
  });

  it("reset() rewinds the click cursor and clears prepared styles", async () => {
    const t1: FakeTarget = { cNvPrId: 11, style: {}, animations: [], pendingResolve: null };
    const slide = makeSlide([fadeAnimation(11, 1, "onClick")]);
    const root = makeRoot([t1]);
    const ctrl = createPlayback(root, slide, { slideSize: SLIDE_SIZE });
    ctrl.prepare();
    await ctrl.clickAdvance();
    expect(ctrl.hasMore()).toBe(false);

    ctrl.reset();
    // Styles wiped (the engine restores the empty inline default so
    // SVG inheritance / class styling can take over again).
    expect(t1.style.visibility).toBe("");
    expect(t1.style.opacity).toBe("");
    // Cursor rewound — the previously-played group is replayable.
    expect(ctrl.hasMore()).toBe(true);
    ctrl.destroy();
  });
});
