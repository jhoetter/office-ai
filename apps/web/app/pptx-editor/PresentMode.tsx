"use client";

import * as React from "react";
import { ChevronLeft, ChevronRight, Maximize2, Minimize2, Monitor, X } from "lucide-react";
import type { ChartPart, NotesSlide, PptxSnapshot, Slide, SlideSize, ThemeColorScheme } from "@officeai/pptx";
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
export interface PresentModeProps {
  readonly snapshot: PptxSnapshot;
  readonly initialSlideIndex: number;
  readonly mediaUrls?: ReadonlyMap<string, string>;
  readonly charts?: ReadonlyMap<string, ChartPart>;
  readonly onClose: () => void;
}

export function PresentMode(props: PresentModeProps): React.ReactElement {
  const slides = props.snapshot.root.slides;
  const slideSize = props.snapshot.root.slideSize;
  const theme = props.snapshot.root.themeDefault;
  const notesSlides = props.snapshot.root.notesSlides;

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
  const next = React.useCallback(() => setIndex((i) => Math.min(i + 1, total - 1)), [total]);
  const prev = React.useCallback(() => setIndex((i) => Math.max(i - 1, 0)), []);
  const first = React.useCallback(() => setIndex(0), []);
  const last = React.useCallback(() => setIndex(total - 1), [total]);

  const close = React.useCallback(() => {
    if (document.fullscreenElement) void document.exitFullscreen().catch(() => undefined);
    props.onClose();
  }, [props]);

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
          next();
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
        {/* Stage */}
        <button
          type="button"
          onClick={next}
          className="relative flex min-h-0 flex-1 cursor-pointer items-center justify-center bg-transparent"
          aria-label="Advance slide"
          data-testid="pptx-present-stage"
        >
          {currentSlide ? (
            <SlideStage
              slide={currentSlide}
              slideSize={slideSize}
              mediaUrls={props.mediaUrls}
              theme={theme}
              charts={props.charts}
              aspectRatio={aspect}
            />
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
                onClick={next}
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
  // CSS-only `aspect-ratio` + `max-width/max-height: 100%` collapses to
  // 0×0 inside a flex parent because the inline SVG carries no
  // intrinsic size (only viewBox), so neither axis ever resolves to a
  // definite length. Measure the available area and compute the
  // largest aspect-preserving rect that fits — this is what made the
  // present-mode stage render as a solid black screen.
  const fitRef = React.useRef<HTMLDivElement | null>(null);
  const [size, setSize] = React.useState<{ width: number; height: number } | null>(null);
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
        prev && Math.abs(prev.width - w) < 0.5 && Math.abs(prev.height - h) < 0.5 ? prev : { width: w, height: h }
      );
    };
    update();
    if (typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver(update);
    ro.observe(el);
    return () => ro.disconnect();
  }, [props.aspectRatio]);
  return (
    <div
      ref={fitRef}
      className="flex h-full w-full items-center justify-center"
    >
      {size ? (
        <div
          className="bg-white shadow-2xl"
          style={{ width: size.width, height: size.height }}
          dangerouslySetInnerHTML={{ __html: svg }}
        />
      ) : null}
    </div>
  );
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
