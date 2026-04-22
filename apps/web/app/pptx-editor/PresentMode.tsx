"use client";

import * as React from "react";
import { ChevronLeft, ChevronRight, Maximize2, Minimize2, Monitor, X } from "lucide-react";
import {
  createPlayback,
  type ChartPart,
  type MediaShape,
  type NotesSlide,
  type PlaybackController,
  type PptxSnapshot,
  type Shape,
  type Slide,
  type SlideSize,
  type SlideTransition,
  type ThemeColorScheme,
} from "@officeai/pptx";
import { slideAspectRatio, slideToSvgString } from "@officeai/pptx/renderer";

/**
 * D10 — Present mode + speaker view.
 *
 * Full-screen, non-interactive playback of the current snapshot. We
 * reuse `slideToSvgString` so what users see during a rehearsal is
 * exactly what the parser/serializer round-trip produces — no second
 * renderer to drift out of sync. Speaker view is a side panel toggled
 * by `S` that surfaces the active slide's speaker notes plus a small
 * preview of the next slide and a wall-clock + elapsed timer.
 *
 * Keyboard model (mirrors PowerPoint as closely as a single-screen
 * web app reasonably can):
 *   →  ↓  Space  PageDown    next slide
 *   ←  ↑          PageUp     previous slide
 *   Home / End                first / last slide
 *   S                         toggle speaker view
 *   F                         toggle browser fullscreen
 *   Esc                       leave present mode
 */
/**
 * Minimal "snapshot source" interface so PresentMode can re-read the
 * live snapshot without taking a hard dependency on `PptxAgent`. The
 * editor wires the live agent in; tests / standalone callers can pass
 * a static snapshot through `snapshot` and skip `subscribeSnapshot`.
 */
export interface SnapshotSource {
  readonly getSnapshot: () => PptxSnapshot;
  readonly subscribe: (listener: () => void) => () => void;
}

export interface PresentModeProps {
  /**
   * Initial snapshot. Used as the first frame; if `subscribeSnapshot`
   * is also provided, subsequent renders pull from there so a
   * just-applied transition / animation command shows up immediately
   * instead of being silently swallowed by a stale prop.
   */
  readonly snapshot: PptxSnapshot;
  /**
   * Optional live source. When present, PresentMode subscribes to it
   * on mount and always renders against the freshest snapshot — this
   * fixes the "I assigned Fade then opened presenter and nothing
   * played" race where `setPresenting(true)` and `applyCommand` raced
   * for the same React batch.
   */
  readonly subscribeSnapshot?: SnapshotSource;
  readonly initialSlideIndex: number;
  readonly mediaUrls?: ReadonlyMap<string, string>;
  readonly charts?: ReadonlyMap<string, ChartPart>;
  readonly onClose: () => void;
}

export function PresentMode(props: PresentModeProps): React.ReactElement {
  // Live snapshot. When the parent supplies `subscribeSnapshot`, we
  // ignore subsequent `props.snapshot` updates and pull from the
  // source ourselves — this guarantees we never render against a
  // snapshot that pre-dates our mount frame.
  const [liveSnapshot, setLiveSnapshot] = React.useState<PptxSnapshot>(
    () => props.subscribeSnapshot?.getSnapshot() ?? props.snapshot
  );
  React.useEffect(() => {
    const src = props.subscribeSnapshot;
    if (!src) return;
    setLiveSnapshot(src.getSnapshot());
    const unsub = src.subscribe(() => setLiveSnapshot(src.getSnapshot()));
    return () => unsub();
  }, [props.subscribeSnapshot]);
  const snapshot = props.subscribeSnapshot ? liveSnapshot : props.snapshot;
  const slides = snapshot.root.slides;
  const slideSize = snapshot.root.slideSize;
  const theme = snapshot.root.themeDefault;
  const notesSlides = snapshot.root.notesSlides;

  const [index, setIndex] = React.useState(() => clampIndex(props.initialSlideIndex, slides.length));
  const [showSpeaker, setShowSpeaker] = React.useState(false);
  const [isFullscreen, setIsFullscreen] = React.useState(false);
  const [startedAt] = React.useState(() => Date.now());
  const [now, setNow] = React.useState(() => Date.now());
  const containerRef = React.useRef<HTMLDivElement | null>(null);

  React.useEffect(() => {
    const id = window.setInterval(() => setNow(Date.now()), 1_000);
    return () => window.clearInterval(id);
  }, []);

  React.useEffect(() => {
    const onChange = (): void => setIsFullscreen(document.fullscreenElement === containerRef.current);
    document.addEventListener("fullscreenchange", onChange);
    return () => document.removeEventListener("fullscreenchange", onChange);
  }, []);

  const total = slides.length;
  // F4 Phase 4 — present-mode playback wiring. The active slide's
  // SlideStage attaches its SVG to this ref so we can build/tear down
  // a `PlaybackController` per slide. The controller hides shapes that
  // carry an entrance animation, then advances one click-group per
  // forward keypress until exhausted; only then does the slide advance
  // to the next one — matching PowerPoint's behaviour.
  const playbackRef = React.useRef<PlaybackController | null>(null);
  // Mirror the controller's `hasMore()` into React state so the footer
  // hint can show a live "click to advance" counter without polling
  // the controller every frame.
  const [pendingAnims, setPendingAnims] = React.useState(0);
  const advanceLockRef = React.useRef(false);
  const advanceSlide = React.useCallback(() => setIndex((i) => Math.min(i + 1, total - 1)), [total]);
  const next = React.useCallback(async () => {
    if (advanceLockRef.current) return;
    const controller = playbackRef.current;
    if (controller && controller.hasMore()) {
      advanceLockRef.current = true;
      try {
        await controller.clickAdvance();
      } finally {
        advanceLockRef.current = false;
        setPendingAnims(controller.hasMore() ? countRemainingGroups(controller) : 0);
      }
      return;
    }
    advanceSlide();
  }, [advanceSlide]);
  const prev = React.useCallback(() => setIndex((i) => Math.max(i - 1, 0)), []);
  const first = React.useCallback(() => setIndex(0), []);
  const last = React.useCallback(() => setIndex(total - 1), [total]);

  // When the slide index changes, drop the previous controller so the
  // new SlideStage can wire up its own. SlideStage owns the create()
  // call (it has the SVG ref) and reports back via `onPlaybackReady`.
  React.useEffect(() => {
    return () => {
      playbackRef.current?.reset();
      playbackRef.current?.destroy();
      playbackRef.current = null;
    };
  }, [index]);

  const handlePlaybackReady = React.useCallback((controller: PlaybackController | null) => {
    if (playbackRef.current && playbackRef.current !== controller) {
      playbackRef.current.reset();
      playbackRef.current.destroy();
    }
    playbackRef.current = controller;
    setPendingAnims(controller && controller.hasMore() ? countRemainingGroups(controller) : 0);
  }, []);

  const close = React.useCallback(() => {
    if (document.fullscreenElement) void document.exitFullscreen().catch(() => undefined);
    props.onClose();
  }, [props.onClose]);

  const toggleFullscreen = React.useCallback(() => {
    const el = containerRef.current;
    if (!el) return;
    if (document.fullscreenElement === el) {
      void document.exitFullscreen().catch(() => undefined);
    } else {
      void el.requestFullscreen().catch(() => undefined);
    }
  }, []);

  React.useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      // Don't interfere with text input — there shouldn't be any in
      // present mode, but speaker-notes textareas (future) might.
      const target = e.target as HTMLElement | null;
      if (
        target &&
        (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)
      ) {
        return;
      }
      switch (e.key) {
        case "ArrowRight":
        case "ArrowDown":
        case "PageDown":
        case " ":
          e.preventDefault();
          void next();
          return;
        case "ArrowLeft":
        case "ArrowUp":
        case "PageUp":
          e.preventDefault();
          prev();
          return;
        case "Home":
          e.preventDefault();
          first();
          return;
        case "End":
          e.preventDefault();
          last();
          return;
        case "Escape":
          e.preventDefault();
          close();
          return;
        case "s":
        case "S":
          e.preventDefault();
          setShowSpeaker((v) => !v);
          return;
        case "f":
        case "F":
          e.preventDefault();
          toggleFullscreen();
          return;
        default:
          return;
      }
    };
    window.addEventListener("keydown", onKey, { capture: true });
    return () => window.removeEventListener("keydown", onKey, { capture: true } as EventListenerOptions);
  }, [close, first, last, next, prev, toggleFullscreen]);

  const currentSlide = slides[index];
  const nextSlide = slides[index + 1] ?? null;

  // F-C2: slide-to-slide transitions. Keep a snapshot of the previous
  // slide whenever `index` advances forward (Office's transition is
  // applied on the **incoming** slide; backwards navigation is
  // intentionally instant — PowerPoint's behaviour). The overlay
  // mounts both slides, runs WAAPI on the matching motion, then
  // unmounts the previous one.
  //
  // Initialise the ref to a sentinel (-1) so the first effect pass
  // sees `prev !== index` and the **incoming-slide transition fires
  // on mount**. PowerPoint plays the entry transition for the slide
  // the user lands on; the previous behaviour (initialise to `index`)
  // silently swallowed it.
  const prevIndexRef = React.useRef<number>(-1);
  const [transitioning, setTransitioning] = React.useState<{
    fromIndex: number;
    toIndex: number;
    kind: SlideTransition["kind"];
    speed: SlideTransition["speed"] | undefined;
  } | null>(null);
  React.useEffect(() => {
    const prev = prevIndexRef.current;
    if (prev === index) return;
    if (index > prev) {
      const t = currentSlide?.transition;
      if (t && t.kind !== "none" && t.kind !== "unsupported") {
        setTransitioning({ fromIndex: prev, toIndex: index, kind: t.kind, speed: t.speed });
      }
    }
    prevIndexRef.current = index;
  }, [index, currentSlide]);
  const handleTransitionDone = React.useCallback(() => setTransitioning(null), []);
  const notesText = currentSlide ? readNotesText(notesSlides, currentSlide.notesSlidePartPath) : "";

  const aspect = slideAspectRatio(slideSize);

  return (
    <div
      ref={containerRef}
      data-testid="pptx-present-mode"
      role="dialog"
      aria-modal="true"
      aria-label="Presentation"
      className="fixed inset-0 z-[100] flex flex-col bg-black text-white"
    >
      <header className="flex items-center justify-between gap-3 px-3 py-2 text-xs">
        <div className="flex items-center gap-2 opacity-70">
          <span data-testid="pptx-present-counter">
            Slide {index + 1} / {total}
          </span>
          <span className="hidden sm:inline">·</span>
          <span className="hidden sm:inline">Elapsed {formatElapsed(now - startedAt)}</span>
          {pendingAnims > 0 ? (
            <>
              <span className="hidden sm:inline">·</span>
              <span
                data-testid="pptx-present-pending-anims"
                className="rounded bg-white/15 px-1.5 py-0.5 text-[10px] uppercase tracking-wide text-white"
              >
                Animations pending — click to play
              </span>
            </>
          ) : null}
          {currentSlide?.transition &&
          currentSlide.transition.kind !== "none" &&
          currentSlide.transition.kind !== "unsupported" ? (
            <>
              <span className="hidden sm:inline">·</span>
              <span
                data-testid="pptx-present-transition"
                className="rounded bg-white/10 px-1.5 py-0.5 text-[10px] uppercase tracking-wide text-white/80"
                title={`Transition: ${currentSlide.transition.kind}`}
              >
                {currentSlide.transition.kind}
              </span>
            </>
          ) : null}
        </div>
        <div className="flex items-center gap-1">
          <PresentBarButton
            label="Speaker view (S)"
            active={showSpeaker}
            onClick={() => setShowSpeaker((v) => !v)}
            icon={<Monitor size={14} />}
          />
          <PresentBarButton
            label={isFullscreen ? "Exit fullscreen (F)" : "Fullscreen (F)"}
            onClick={toggleFullscreen}
            icon={isFullscreen ? <Minimize2 size={14} /> : <Maximize2 size={14} />}
          />
          <PresentBarButton
            label="Exit present mode (Esc)"
            onClick={close}
            icon={<X size={14} />}
            testId="pptx-present-close"
          />
        </div>
      </header>

      <div className="flex min-h-0 flex-1 gap-3 px-3 pb-3">
        {/* Stage. We pulse the outline on the very first slide whenever
         * something is waiting for a click — either the slide carries
         * a transition kind that hasn't played yet, or the SlideStage
         * has reported entrance animations pending. PowerPoint shows
         * a similar hint via the cursor; on a single-screen web app
         * the user expects a clearer visual cue. */}
        <button
          type="button"
          onClick={() => void next()}
          className={`relative flex min-h-0 flex-1 cursor-pointer items-center justify-center bg-transparent ${
            index === 0 &&
            (pendingAnims > 0 ||
              (currentSlide?.transition &&
                currentSlide.transition.kind !== "none" &&
                currentSlide.transition.kind !== "unsupported"))
              ? "ring-2 ring-white/30 [animation:pptx-pulse_1.6s_ease-in-out_infinite]"
              : ""
          }`}
          aria-label="Advance slide"
          data-testid="pptx-present-stage"
          data-pulse={
            index === 0 &&
            (pendingAnims > 0 ||
              (currentSlide?.transition &&
                currentSlide.transition.kind !== "none" &&
                currentSlide.transition.kind !== "unsupported"))
              ? "true"
              : "false"
          }
        >
          {currentSlide ? (
            <div className="relative h-full w-full">
              <SlideStage
                slide={currentSlide}
                slideSize={slideSize}
                mediaUrls={props.mediaUrls}
                theme={theme}
                charts={props.charts}
                aspectRatio={aspect}
                onPlaybackReady={handlePlaybackReady}
              />
              {transitioning && transitioning.toIndex === index ? (
                <SlideTransitionOverlay
                  fromSlide={slides[transitioning.fromIndex] ?? null}
                  toSlide={currentSlide}
                  slideSize={slideSize}
                  mediaUrls={props.mediaUrls}
                  theme={theme}
                  charts={props.charts}
                  aspectRatio={aspect}
                  kind={transitioning.kind}
                  speed={transitioning.speed}
                  onDone={handleTransitionDone}
                />
              ) : null}
            </div>
          ) : (
            <span className="text-sm opacity-60">No slides</span>
          )}
        </button>

        {/* Speaker view */}
        {showSpeaker ? (
          <aside
            data-testid="pptx-present-speaker"
            className="flex w-[320px] shrink-0 flex-col gap-3 overflow-hidden rounded-md bg-zinc-900/80 p-3 text-xs"
          >
            <div>
              <div className="mb-1 text-[10px] uppercase tracking-wide text-zinc-400">Speaker notes</div>
              <div
                data-testid="pptx-present-notes"
                className="max-h-[42vh] min-h-[120px] overflow-auto whitespace-pre-wrap rounded border border-zinc-700 bg-zinc-950 p-2 text-sm leading-relaxed"
              >
                {notesText.length > 0 ? (
                  notesText
                ) : (
                  <span className="text-zinc-500">No notes for this slide.</span>
                )}
              </div>
            </div>
            <div>
              <div className="mb-1 text-[10px] uppercase tracking-wide text-zinc-400">Up next</div>
              {nextSlide ? (
                <NextSlidePreview
                  slide={nextSlide}
                  slideSize={slideSize}
                  mediaUrls={props.mediaUrls}
                  theme={theme}
                  charts={props.charts}
                />
              ) : (
                <div className="rounded border border-dashed border-zinc-700 p-4 text-center text-zinc-500">
                  Last slide
                </div>
              )}
            </div>
            <div className="mt-auto flex items-center justify-between gap-2 text-zinc-400">
              <button
                type="button"
                onClick={prev}
                disabled={index === 0}
                className="inline-flex items-center gap-1 rounded px-2 py-1 hover:bg-zinc-800 disabled:opacity-30"
              >
                <ChevronLeft size={14} /> Prev
              </button>
              <span>{formatClock(now)}</span>
              <button
                type="button"
                onClick={() => void next()}
                disabled={index === total - 1}
                className="inline-flex items-center gap-1 rounded px-2 py-1 hover:bg-zinc-800 disabled:opacity-30"
              >
                Next <ChevronRight size={14} />
              </button>
            </div>
          </aside>
        ) : null}
      </div>
    </div>
  );
}

interface SlideStageProps {
  readonly slide: Slide;
  readonly slideSize: SlideSize;
  readonly mediaUrls?: ReadonlyMap<string, string>;
  readonly theme?: ThemeColorScheme;
  readonly charts?: ReadonlyMap<string, ChartPart>;
  readonly aspectRatio: number;
  /**
   * F4 Phase 4 — once the SVG has been mounted, hand a fresh playback
   * controller back to the parent (`PresentMode`) so `Space` / `Arrow` /
   * click can advance through the slide's animations before flipping
   * to the next slide. Receives `null` on unmount so the parent can
   * drop the reference cleanly.
   */
  readonly onPlaybackReady?: (controller: PlaybackController | null) => void;
}

function SlideStage(props: SlideStageProps): React.ReactElement {
  const svg = React.useMemo(
    () =>
      slideToSvgString(props.slide, {
        slideSize: props.slideSize,
        ...(props.mediaUrls ? { mediaUrls: props.mediaUrls } : {}),
        ...(props.theme ? { theme: props.theme } : {}),
        ...(props.charts ? { charts: props.charts } : {}),
      }),
    [props.slide, props.slideSize, props.mediaUrls, props.theme, props.charts]
  );
  const slideHostRef = React.useRef<HTMLDivElement | null>(null);
  // Same `size` measurement that gates the host element below — we
  // hoist it here so the playback effect can react to it. Previously
  // the effect ran with a `null` host (because the host was only
  // mounted *after* `size` resolved on a later frame) and the early
  // bail at `if (!host)` silently swallowed every animation. The fix
  // is twofold: (1) include `size` in the dep list so the effect
  // re-runs when the host appears; (2) short-circuit when `size` is
  // null so we don't even try to query the not-yet-mounted SVG.
  const fitRef = React.useRef<HTMLDivElement | null>(null);
  const [size, setSize] = React.useState<{ width: number; height: number } | null>(null);
  React.useEffect(() => {
    const host = slideHostRef.current;
    const onReady = props.onPlaybackReady;
    if (!host || !onReady || !size) return;
    const svgEl = host.querySelector<SVGSVGElement>("svg");
    if (!svgEl) return;
    const controller = createPlayback(svgEl, props.slide, {
      slideSize: props.slideSize,
    });
    controller.prepare();
    if (typeof window !== "undefined" && window.location.search.includes("debug=anim")) {
      console.debug(
        "[playback] slide=%s groups=%d hasMore=%s",
        props.slide.id,
        countRemainingGroups(controller),
        String(controller.hasMore())
      );
    }
    onReady(controller);
    return () => {
      onReady(null);
      controller.reset();
      controller.destroy();
    };
    // `svg` is the rendered HTML string; when it changes we have a
    // new SVG element to attach to. `size` belongs in the dep list
    // because the SVG host element is only mounted *after* size
    // resolves on a later frame; without it the very first effect
    // pass sees `host === null` and silently bails (which was the
    // animation playback regression).
  }, [props.slide, props.slideSize, props.onPlaybackReady, svg, size]);
  // CSS-only `aspect-ratio` + `max-width/max-height: 100%` collapses to
  // 0×0 inside a flex parent because the inline SVG carries no
  // intrinsic size (only viewBox), so neither axis ever resolves to a
  // definite length. Measure the available area and compute the
  // largest aspect-preserving rect that fits — this is what made the
  // present-mode stage render as a solid black screen.
  React.useLayoutEffect(() => {
    const el = fitRef.current;
    if (!el) return;
    const update = (): void => {
      const rect = el.getBoundingClientRect();
      if (rect.width <= 0 || rect.height <= 0) return;
      const containerAspect = rect.width / rect.height;
      const w = containerAspect > props.aspectRatio ? rect.height * props.aspectRatio : rect.width;
      const h = containerAspect > props.aspectRatio ? rect.height : rect.width / props.aspectRatio;
      setSize((prev) =>
        prev && Math.abs(prev.width - w) < 0.5 && Math.abs(prev.height - h) < 0.5
          ? prev
          : { width: w, height: h }
      );
    };
    update();
    if (typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver(update);
    ro.observe(el);
    return () => ro.disconnect();
  }, [props.aspectRatio]);
  // Collect every `MediaShape` on the slide so Present mode can mount
  // a real `<video>` / `<audio>` element on top of the static SVG.
  // The renderer paints a placeholder rectangle for media — fine for
  // edit-mode preview / export, but in a rehearsal the user expects
  // playable controls. We resolve absolute slide-coords per shape so
  // shapes nested inside a group still land at the right spot.
  const mediaShapes = React.useMemo(() => collectMediaShapes(props.slide.shapes), [props.slide]);
  return (
    <div ref={fitRef} className="flex h-full w-full items-center justify-center">
      {size ? (
        <div
          ref={slideHostRef}
          className="relative bg-white shadow-2xl"
          style={{ width: size.width, height: size.height }}
        >
          <div className="absolute inset-0" dangerouslySetInnerHTML={{ __html: svg }} />
          {mediaShapes.map(({ shape, xEmu, yEmu }) => {
            const url = props.mediaUrls?.get(shape.mediaPath);
            if (!url || !shape.size) return null;
            const pct = {
              left: `${(xEmu / props.slideSize.cxEmu) * 100}%`,
              top: `${(yEmu / props.slideSize.cyEmu) * 100}%`,
              width: `${(shape.size.cxEmu / props.slideSize.cxEmu) * 100}%`,
              height: `${(shape.size.cyEmu / props.slideSize.cyEmu) * 100}%`,
            };
            const common = {
              src: url,
              controls: true,
              preload: "metadata" as const,
              "data-testid": `pptx-present-media-${shape.id}`,
            };
            return (
              <div key={shape.id} className="absolute" style={pct}>
                {shape.mediaType === "video" ? (
                  <video {...common} className="h-full w-full bg-black object-contain" />
                ) : (
                  <audio {...common} className="h-full w-full" />
                )}
              </div>
            );
          })}
        </div>
      ) : null}
    </div>
  );
}

interface PositionedMedia {
  readonly shape: MediaShape;
  /** Absolute slide-coord top-left after walking through any group ancestry. */
  readonly xEmu: number;
  readonly yEmu: number;
}

/**
 * Walk the slide tree and surface every `MediaShape` together with its
 * resolved absolute slide-coord position. Group children carry coords
 * relative to the group origin (matching the `<a:chOff>` offset the
 * SVG renderer applies via a `<g transform="translate(...)">`); we
 * accumulate the parent group's position so a media shape inside a
 * group still lands at the right spot in the HTML overlay.
 */
function collectMediaShapes(shapes: ReadonlyArray<Shape>): ReadonlyArray<PositionedMedia> {
  const out: PositionedMedia[] = [];
  const visit = (list: ReadonlyArray<Shape>, dx: number, dy: number): void => {
    for (const s of list) {
      if (s.kind === "media") {
        if (s.position) {
          out.push({ shape: s, xEmu: s.position.xEmu + dx, yEmu: s.position.yEmu + dy });
        }
      } else if (s.kind === "group") {
        const gx = (s.position?.xEmu ?? 0) + dx;
        const gy = (s.position?.yEmu ?? 0) + dy;
        visit(s.children, gx, gy);
      }
    }
  };
  visit(shapes, 0, 0);
  return out;
}

interface NextSlidePreviewProps {
  readonly slide: Slide;
  readonly slideSize: SlideSize;
  readonly mediaUrls?: ReadonlyMap<string, string>;
  readonly theme?: ThemeColorScheme;
  readonly charts?: ReadonlyMap<string, ChartPart>;
}

function NextSlidePreview(props: NextSlidePreviewProps): React.ReactElement {
  const svg = React.useMemo(
    () =>
      slideToSvgString(props.slide, {
        slideSize: props.slideSize,
        ...(props.mediaUrls ? { mediaUrls: props.mediaUrls } : {}),
        ...(props.theme ? { theme: props.theme } : {}),
        ...(props.charts ? { charts: props.charts } : {}),
      }),
    [props.slide, props.slideSize, props.mediaUrls, props.theme, props.charts]
  );
  const aspect = slideAspectRatio(props.slideSize);
  return (
    <div
      className="overflow-hidden rounded border border-zinc-700 bg-white"
      style={{ aspectRatio: String(aspect), width: "100%" }}
      dangerouslySetInnerHTML={{ __html: svg }}
    />
  );
}

interface PresentBarButtonProps {
  readonly label: string;
  readonly icon: React.ReactNode;
  readonly onClick: () => void;
  readonly active?: boolean;
  readonly testId?: string;
}

function PresentBarButton(props: PresentBarButtonProps): React.ReactElement {
  return (
    <button
      type="button"
      onClick={props.onClick}
      title={props.label}
      aria-label={props.label}
      {...(props.testId ? { "data-testid": props.testId } : {})}
      className={[
        "inline-flex items-center gap-1 rounded px-2 py-1 text-xs",
        props.active ? "bg-white/15" : "hover:bg-white/10",
      ].join(" ")}
    >
      {props.icon}
    </button>
  );
}

function readNotesText(notesSlides: ReadonlyMap<string, NotesSlide>, partPath: string | undefined): string {
  if (!partPath) return "";
  const notes = notesSlides.get(partPath);
  if (!notes) return "";
  return notes.body.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n");
}

function clampIndex(i: number, total: number): number {
  if (total <= 0) return 0;
  if (i < 0) return 0;
  if (i >= total) return total - 1;
  return i;
}

function formatElapsed(ms: number): string {
  const totalSec = Math.max(0, Math.floor(ms / 1000));
  const h = Math.floor(totalSec / 3600);
  const m = Math.floor((totalSec % 3600) / 60);
  const s = totalSec % 60;
  if (h > 0) {
    return `${h}:${pad2(m)}:${pad2(s)}`;
  }
  return `${m}:${pad2(s)}`;
}

function formatClock(ms: number): string {
  const d = new Date(ms);
  return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
}

function pad2(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

/**
 * Best-effort live count of click-groups still pending on the active
 * `PlaybackController`. The controller doesn't expose its internal
 * cursor (intentional — its API is verb-only), but the only stable
 * signal we need is "is there at least one more group?". We expose
 * that as `1+` so the footer hint can read "More animations — click to
 * play". Tests don't depend on the precise number.
 */
function countRemainingGroups(controller: PlaybackController): number {
  return controller.hasMore() ? 1 : 0;
}

// ─── F-C2: slide-to-slide transitions ────────────────────────────────

interface SlideTransitionOverlayProps {
  /**
   * The slide being animated *off* the stage. `null` when the
   * transition is the **initial-mount entry** of `toSlide` — in that
   * case the overlay paints a solid backdrop that gets animated out
   * to reveal `toSlide` underneath, so the user actually sees the
   * transition that was assigned to the landed-on slide.
   */
  readonly fromSlide: Slide | null;
  readonly toSlide: Slide;
  readonly slideSize: SlideSize;
  readonly mediaUrls?: ReadonlyMap<string, string>;
  readonly theme?: ThemeColorScheme;
  readonly charts?: ReadonlyMap<string, ChartPart>;
  readonly aspectRatio: number;
  readonly kind: SlideTransition["kind"];
  readonly speed: SlideTransition["speed"] | undefined;
  readonly onDone: () => void;
}

/**
 * Render the previous slide on top of the next, then animate it off
 * via the WAAPI motion appropriate to the captured transition kind.
 * The overlay drops itself when the animation finishes (or after a
 * safety timeout in case WAAPI never fires `finish`).
 *
 * Only the four most-common kinds get bespoke motions: fade, push,
 * wipe, split. `cut` resolves to instant. Anything else falls back
 * to a quick crossfade so the transition is visible but not jarring.
 */
function SlideTransitionOverlay(props: SlideTransitionOverlayProps): React.ReactElement | null {
  const durationMs = transitionDuration(props.speed);
  const fromSvg = React.useMemo(
    () =>
      props.fromSlide
        ? slideToSvgString(props.fromSlide, {
            slideSize: props.slideSize,
            ...(props.mediaUrls ? { mediaUrls: props.mediaUrls } : {}),
            ...(props.theme ? { theme: props.theme } : {}),
            ...(props.charts ? { charts: props.charts } : {}),
          })
        : null,
    [props.fromSlide, props.slideSize, props.mediaUrls, props.theme, props.charts]
  );
  const overlayRef = React.useRef<HTMLDivElement | null>(null);
  // Capture the latest `onDone` in a ref so the playback effect can
  // depend only on the stable kind / duration values. Previously the
  // effect's dep list was `[props, durationMs]`, which re-ran on every
  // parent render (the `setNow` second-tick alone is enough). Each
  // re-run cancelled the in-flight animation; in React Strict Mode the
  // double-mount cycle cancelled it before the very first frame, which
  // is why "I assigned Fade then opened presenter and nothing played"
  // resurfaced even after the snapshot-source fix. The overlay mounts
  // once per transition, so a ref-captured callback is the right shape.
  const onDoneRef = React.useRef(props.onDone);
  onDoneRef.current = props.onDone;
  const kind = props.kind;
  React.useEffect(() => {
    // Track whether this effect instance has been torn down so the
    // `cancel` event we fire in cleanup doesn't bubble out to onDone
    // and unmount the overlay. Critical in React Strict Mode where the
    // first cleanup fires before the second mount runs — without this
    // guard, the cancel-handler short-circuits the transition before a
    // single keyframe paints.
    let teardown = false;
    const fireDone = (): void => {
      if (teardown) return;
      onDoneRef.current();
    };
    if (kind === "cut") {
      fireDone();
      return;
    }
    const el = overlayRef.current;
    if (!el) {
      fireDone();
      return;
    }
    const animation = playTransitionAnimation(el, kind, durationMs);
    const safetyId = window.setTimeout(fireDone, durationMs + 200);
    if (animation) {
      animation.addEventListener("finish", fireDone, { once: true });
      animation.addEventListener("cancel", fireDone, { once: true });
    }
    return () => {
      teardown = true;
      window.clearTimeout(safetyId);
      animation?.cancel();
    };
  }, [kind, durationMs]);
  if (props.kind === "cut") return null;
  return (
    <div
      data-testid="pptx-present-transition-overlay"
      data-kind={props.kind}
      className="pointer-events-none absolute inset-0 flex items-center justify-center"
    >
      <div
        ref={overlayRef}
        className="bg-white shadow-2xl will-change-transform"
        style={{
          aspectRatio: String(props.aspectRatio),
          maxWidth: "100%",
          maxHeight: "100%",
          width: "100%",
          height: "100%",
        }}
        // When `fromSlide` is null this is the entry transition for
        // the landed-on slide. A blank white backdrop animates out to
        // reveal the slide underneath — same WAAPI keyframes, just
        // no inner SVG to draw.
        {...(fromSvg !== null ? { dangerouslySetInnerHTML: { __html: fromSvg } } : {})}
      />
    </div>
  );
}

function transitionDuration(speed: SlideTransition["speed"] | undefined): number {
  switch (speed) {
    case "fast":
      return 250;
    case "slow":
      return 750;
    case "med":
    default:
      return 500;
  }
}

function playTransitionAnimation(
  el: HTMLElement,
  kind: SlideTransition["kind"],
  durationMs: number
): Animation | null {
  // Delegated per-kind keyframes. WAAPI returns null in non-browser
  // jest envs; callers handle that by short-circuiting via `onDone`.
  const easing = "cubic-bezier(0.4, 0, 0.2, 1)";
  switch (kind) {
    case "fade":
      return el.animate([{ opacity: 1 }, { opacity: 0 }], { duration: durationMs, easing, fill: "forwards" });
    case "push":
      // PowerPoint's default push direction is "left" (incoming from
      // right). We translate the outgoing slide off to the left.
      return el.animate([{ transform: "translateX(0%)" }, { transform: "translateX(-100%)" }], {
        duration: durationMs,
        easing,
        fill: "forwards",
      });
    case "wipe":
      return el.animate([{ clipPath: "inset(0 0 0 0)" }, { clipPath: "inset(0 100% 0 0)" }], {
        duration: durationMs,
        easing,
        fill: "forwards",
      });
    case "split":
      return el.animate([{ clipPath: "inset(0 0 0 0)" }, { clipPath: "inset(0 50% 0 50%)" }], {
        duration: durationMs,
        easing,
        fill: "forwards",
      });
    case "none":
    case "cut":
    case "unsupported":
    default:
      // Fallback crossfade so the user isn't left looking at a frozen
      // overlay if we ever land here.
      return el.animate([{ opacity: 1 }, { opacity: 0 }], {
        duration: Math.min(180, durationMs),
        easing,
        fill: "forwards",
      });
  }
}
