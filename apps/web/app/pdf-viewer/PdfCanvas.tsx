"use client";

import * as React from "react";
import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from "react";
import type { PdfEngineDocument, PdfEnginePage, PdfEngineTextItem } from "@officeai/pdf-engine";
import type {
  AddAnnotationPayload,
  PdfAnnotation,
  PdfPage,
  PdfRect,
  PdfRotation,
  PdfSnapshot,
} from "@officeai/pdf";
import { useTranslator } from "@/lib/i18n";
import { darkModeCssFilter } from "./darkMode";
import type { PdfAnnotationTool } from "./PdfToolbar";

/** Canvas-side view modes — 1:1 with the toolbar enum. */
export type PdfViewMode = "single" | "continuous" | "two-up";

/**
 * Binary dark-mode toggle. CSS-filter based, no per-pixel pass.
 */
export type PdfDarkModeStrategy = "off" | "on";

export interface PdfHighlight {
  readonly pageNumber: number;
  /** Normalized rect [x1, y1, x2, y2] (0..1, origin top-left). */
  readonly rect: readonly [number, number, number, number];
  /** Bumped by the caller to re-trigger the flash animation. */
  readonly nonce: number;
}

export interface PdfCanvasProps {
  readonly engineDoc: PdfEngineDocument | null;
  readonly snapshot: PdfSnapshot | null;
  readonly currentPage: number;
  /**
   * Multiplier on top of the fit-width baseline. 1 = fit width.
   * "Actual size" sets this to a CSS-pixel-equivalent scale (1 PDF
   * unit = 1 CSS pixel); "fit page" computes a value such that the
   * page's content rect (after rotation) fits the viewport on both
   * axes.
   */
  readonly zoom: number;
  readonly viewMode: PdfViewMode;
  readonly darkMode: PdfDarkModeStrategy;
  /**
   * Reflow renders text-only single-column markdown reflow. Falls
   * back to a "scanned page — no text" hint when a page has no
   * text layer.
   */
  readonly reflow: boolean;
  /**
   * Extra clockwise rotation, in 90° increments, applied on top of
   * each page's intrinsic rotation. Independent from the document-
   * mutating `pdf:rotate-pages` command.
   */
  readonly viewportRotation: PdfRotation;
  /**
   * Fired when the user clicks a page or scrolls a different page
   * into the centre of the viewport. The editor uses it to keep
   * its `currentPage` state and the realtime presence broadcast
   * in sync.
   */
  readonly onCurrentPageChange: (pageNumber: number) => void;
  /**
   * Optional highlight rect in normalized page coordinates. The
   * editor's find-replace integration sets this to flash the active
   * search match. Cleared by passing `null`.
   */
  readonly highlight?: PdfHighlight | null;
  /**
   * Reports container-derived zoom metrics back to the editor so
   * the toolbar's "Actual size" and "Fit page" callbacks can pick
   * the right `zoom` value. `actualSizeZoom` makes 1 PDF unit equal
   * 1 CSS pixel; `fitPageZoom` shrinks the largest page to fit the
   * viewport on both axes.
   */
  readonly onZoomMetricsChange?: (metrics: PdfZoomMetrics) => void;
  /**
   * Bumped every time the editor wants the canvas to deliberately
   * scroll to `currentPage`. The canvas only fires its programmatic
   * `scrollTo` when this nonce changes — purely observation-driven
   * page changes (the IO callback) do *not* bump it, so the
   * canvas never fights the user's wheel.
   */
  readonly jumpNonce?: number;
  /**
   * The currently armed annotation tool, or `null`. Controls the
   * page cursor and which canvas-level handlers fire.
   */
  readonly armedTool?: PdfAnnotationTool | null;
  /**
   * Fired when the canvas wants to commit a new annotation
   * (e.g. user released a text selection while the highlight
   * tool was armed).
   */
  readonly onAddAnnotation?: (input: AddAnnotationPayload) => void;
  /**
   * Fired after the canvas successfully commits a new annotation,
   * giving the editor a chance to un-arm the tool.
   */
  readonly onAnnotationCreated?: () => void;
}

/**
 * Container-derived zoom values. The editor uses these so toolbar
 * presets ("Actual size", "Fit page") map to a real `zoom` instead
 * of hard-coded numbers.
 */
export interface PdfZoomMetrics {
  readonly actualSizeZoom: number;
  readonly fitPageZoom: number;
}

/**
 * The PDF rendering surface.
 *
 * Layout: a vertical scroller. Each page is a placeholder `<div>`
 * sized to the page's intrinsic aspect ratio so the scrollbar
 * reflects the full document length immediately, no re-layout once
 * lazy renders resolve. An IntersectionObserver mounts each
 * page's `<canvas>` only when its placeholder enters the viewport
 * (with a 1-page rootMargin so the next page is ready before the
 * user reaches it).
 *
 * Engine-agnostic: only consumes the `PdfEngineDocument` interface,
 * never imports the PDF.js backend directly. Worker setup happens
 * at editor scope; by the time we receive `engineDoc`, the worker
 * is wired and `getPage()` works.
 */
export function PdfCanvas(props: PdfCanvasProps): ReactNode {
  const {
    engineDoc,
    snapshot,
    currentPage,
    zoom,
    viewMode,
    darkMode,
    reflow,
    viewportRotation,
    onCurrentPageChange,
    highlight,
    onZoomMetricsChange,
    jumpNonce,
    armedTool = null,
    onAddAnnotation,
    onAnnotationCreated,
  } = props;
  const { t } = useTranslator();
  const containerRef = useRef<HTMLDivElement | null>(null);
  const [containerWidth, setContainerWidth] = useState(0);
  const [containerHeight, setContainerHeight] = useState(0);

  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => {
      setContainerWidth(el.clientWidth);
      setContainerHeight(el.clientHeight);
    });
    ro.observe(el);
    setContainerWidth(el.clientWidth);
    setContainerHeight(el.clientHeight);
    return () => ro.disconnect();
  }, []);

  const pages = snapshot?.root.pages ?? [];
  const baseScale = useMemo(() => {
    if (!pages.length || containerWidth === 0) return 1;
    let widest = 0;
    for (const p of pages) {
      const w = effectiveWidth(p, viewportRotation);
      if (w > widest) widest = w;
    }
    if (widest === 0) return 1;
    // Fit the widest page into the container with a small gutter on
    // each side. The user's `zoom` then multiplies on top.
    const gutter = viewMode === "two-up" ? 24 : 48;
    const target = Math.max(120, containerWidth - gutter);
    return target / widest;
  }, [pages, containerWidth, viewportRotation, viewMode]);
  const scale = baseScale * zoom;

  // Surface fit-page / actual-size targets back to the editor so
  // the toolbar's preset callbacks set a real `zoom` instead of a
  // hard-coded heuristic.
  useEffect(() => {
    if (!onZoomMetricsChange) return;
    if (!pages.length || baseScale === 0) return;
    const tallest = pages.reduce(
      (acc, p) => Math.max(acc, effectiveHeight(p, viewportRotation)),
      0
    );
    const verticalGutter = 32;
    const fitPageScale =
      containerHeight > 0 && tallest > 0
        ? Math.min(
            baseScale,
            Math.max(0.05, (containerHeight - verticalGutter) / tallest)
          )
        : baseScale;
    const actualSizeZoom = 1 / baseScale;
    const fitPageZoom = fitPageScale / baseScale;
    onZoomMetricsChange({ actualSizeZoom, fitPageZoom });
  }, [pages, baseScale, containerHeight, viewportRotation, onZoomMetricsChange]);

  // Track which placeholders are intersected so we know which pages
  // to render. The set is mutable so the IO callback can mutate
  // without triggering an immediate re-render; we tick state to
  // force one render after a batch of intersection changes.
  const visibleRef = useRef<Set<number>>(new Set());
  const [, forceTick] = useState(0);
  const tick = useCallback(() => forceTick((n) => n + 1), []);

  useEffect(() => {
    const root = containerRef.current;
    if (!root) return;
    const io = new IntersectionObserver(
      (entries) => {
        let changed = false;
        for (const e of entries) {
          const num = Number((e.target as HTMLElement).dataset.pageNumber ?? 0);
          if (!num) continue;
          const set = visibleRef.current;
          if (e.isIntersecting) {
            if (!set.has(num)) {
              set.add(num);
              changed = true;
            }
          } else if (set.has(num)) {
            set.delete(num);
            changed = true;
          }
        }
        if (changed) tick();
      },
      {
        root,
        // 1 page worth of margin so we render slightly ahead of
        // scroll. Generous rootMargin avoids the "scroll then wait
        // for blank canvas" flicker on fast wheel input.
        rootMargin: "400px 0px 800px 0px",
        threshold: 0,
      }
    );
    const placeholders = root.querySelectorAll<HTMLElement>("[data-page-number]");
    placeholders.forEach((el) => io.observe(el));
    return () => io.disconnect();
  }, [pages, tick, viewMode, scale]);

  // Track the "current page" by scroll position: the page whose
  // centre is closest to the viewport's centreline wins.
  useEffect(() => {
    const root = containerRef.current;
    if (!root) return;
    let raf = 0;
    const onScroll = (): void => {
      if (raf) return;
      raf = requestAnimationFrame(() => {
        raf = 0;
        const placeholders = Array.from(
          root.querySelectorAll<HTMLElement>("[data-page-number]")
        );
        const viewMid = root.scrollTop + root.clientHeight / 2;
        let nearest: { num: number; dist: number } | null = null;
        for (const el of placeholders) {
          const top = el.offsetTop;
          const bottom = top + el.offsetHeight;
          const mid = (top + bottom) / 2;
          const dist = Math.abs(mid - viewMid);
          const num = Number(el.dataset.pageNumber ?? 0);
          if (!num) continue;
          if (!nearest || dist < nearest.dist) nearest = { num, dist };
        }
        if (nearest && nearest.num !== currentPage) {
          onCurrentPageChange(nearest.num);
        }
      });
    };
    root.addEventListener("scroll", onScroll, { passive: true });
    return () => {
      root.removeEventListener("scroll", onScroll);
      if (raf) cancelAnimationFrame(raf);
    };
  }, [currentPage, onCurrentPageChange, pages.length]);

  // Programmatic scroll when the toolbar / sidebar / shortcuts ask
  // for a different page. Gated on `jumpNonce` so that page changes
  // produced by our own IntersectionObserver (i.e. the user
  // scrolling) never re-trigger this — that's what made the wheel
  // feel snappy / fight the user.
  const lastJumpRef = useRef<number | undefined>(jumpNonce);
  useEffect(() => {
    if (jumpNonce === undefined) return;
    if (lastJumpRef.current === jumpNonce) return;
    lastJumpRef.current = jumpNonce;
    const root = containerRef.current;
    if (!root) return;
    if (viewMode === "single") return;
    const target = root.querySelector<HTMLElement>(
      `[data-page-number="${currentPage}"]`
    );
    if (!target) return;
    const top = target.offsetTop;
    const distance = Math.abs(top - root.scrollTop);
    if (distance < 4) return;
    // Smooth feels great for small jumps; for big jumps (e.g.
    // shift-jumping 50 pages) the animation drags. Snap instantly
    // past five page-heights.
    const tallest = pages.reduce(
      (acc, p) => Math.max(acc, effectiveHeight(p, viewportRotation) * scale),
      1
    );
    const useSmooth = distance < tallest * 5;
    root.scrollTo({ top, behavior: useSmooth ? "smooth" : "auto" });
  }, [jumpNonce, currentPage, viewMode, scale, pages, viewportRotation]);

  if (!snapshot || !engineDoc) {
    return (
      <div
        ref={containerRef}
        className="flex h-full w-full items-center justify-center bg-[var(--page-backdrop)] text-sm text-secondary"
        data-testid="pdf-canvas-empty"
      >
        {t("pdf.loading")}
      </div>
    );
  }

  if (reflow) {
    return (
      <div
        ref={containerRef}
        className="h-full w-full overflow-y-auto bg-[var(--page-backdrop)]"
        data-testid="pdf-canvas-reflow"
      >
        <article className="mx-auto max-w-[68ch] px-6 py-10 text-sm leading-relaxed text-primary">
          {pages.map((p) => (
            <ReflowPage key={p.id} page={p} totalPages={pages.length} />
          ))}
        </article>
      </div>
    );
  }

  const renderTargets = collectRenderTargets(pages, viewMode, currentPage);
  const totalPages = pages.length;

  return (
    <div
      ref={containerRef}
      className="h-full w-full overflow-y-auto bg-[var(--page-backdrop)]"
      data-testid="pdf-canvas"
    >
      <div className="mx-auto flex max-w-full flex-col items-center gap-4 py-6">
        {renderTargets.map((row) => (
          <div
            key={row.key}
            className={row.kind === "two-up" ? "flex items-start gap-3" : ""}
          >
            {row.pages.map((p) => (
              <PdfPageRender
                key={p.id}
                page={p}
                totalPages={totalPages}
                engineDoc={engineDoc}
                scale={scale}
                active={p.pageNumber === currentPage}
                visible={visibleRef.current.has(p.pageNumber)}
                viewportRotation={viewportRotation}
                darkMode={darkMode}
                onClick={() => onCurrentPageChange(p.pageNumber)}
                highlight={
                  highlight && highlight.pageNumber === p.pageNumber ? highlight : null
                }
                annotations={pageAnnotations(snapshot, p.pageNumber)}
                armedTool={armedTool}
                onAddAnnotation={onAddAnnotation}
                onAnnotationCreated={onAnnotationCreated}
              />
            ))}
          </div>
        ))}
      </div>
    </div>
  );
}

interface PdfPageRenderProps {
  readonly page: PdfPage;
  readonly totalPages: number;
  readonly engineDoc: PdfEngineDocument;
  readonly scale: number;
  readonly active: boolean;
  readonly visible: boolean;
  readonly viewportRotation: PdfRotation;
  readonly darkMode: PdfDarkModeStrategy;
  readonly onClick: () => void;
  readonly highlight: PdfHighlight | null;
  readonly annotations: ReadonlyArray<PdfAnnotation>;
  readonly armedTool: PdfAnnotationTool | null;
  readonly onAddAnnotation: ((input: AddAnnotationPayload) => void) | undefined;
  readonly onAnnotationCreated: (() => void) | undefined;
}

/**
 * One placeholder + lazy-mounted canvas for a single page. The
 * placeholder always reserves the right pixel rect so the
 * IntersectionObserver and the scroll-driven "current page"
 * tracker work even before the first render lands. The canvas
 * paints when the page becomes visible; we drive the actual
 * paint via a `useEffect` that owns a cancellation token so an
 * in-flight render can't write into the canvas after the user
 * scrolls past.
 */
function PdfPageRender(props: PdfPageRenderProps): ReactNode {
  const {
    page,
    totalPages,
    engineDoc,
    scale,
    active,
    visible,
    viewportRotation,
    darkMode,
    onClick,
    highlight,
    annotations,
    armedTool,
    onAddAnnotation,
    onAnnotationCreated,
  } = props;
  const highlightArmed = armedTool === "highlight" && viewportRotation === 0;
  const stickyArmed = armedTool === "sticky" && viewportRotation === 0;
  const onHighlightSelection = useCallback(
    (rects: ReadonlyArray<PdfRect>) => {
      if (!onAddAnnotation || rects.length === 0) return;
      const union = unionPdfRects(rects);
      onAddAnnotation({
        kind: "highlight",
        pageNumber: page.pageNumber,
        rect: union,
        quadRects: rects,
        color: { r: 1, g: 0.93, b: 0.16 },
      });
      onAnnotationCreated?.();
    },
    [onAddAnnotation, onAnnotationCreated, page.pageNumber]
  );

  // Pending sticky-note placement. `null` means we're not composing.
  // While set, we render the composer popover; the user types
  // contents and either commits (Enter / Save) or aborts (Esc).
  const [stickyDraft, setStickyDraft] = useState<{ x: number; y: number } | null>(null);
  const onPageClickForSticky = useCallback(
    (e: React.MouseEvent<HTMLDivElement>) => {
      if (!stickyArmed) return;
      const target = e.currentTarget;
      const rect = target.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) return;
      const x = (e.clientX - rect.left) / scale;
      const y = page.height - (e.clientY - rect.top) / scale;
      setStickyDraft({ x, y });
      e.stopPropagation();
    },
    [stickyArmed, scale, page.height]
  );
  const onCommitSticky = useCallback(
    (contents: string) => {
      if (!stickyDraft) return;
      const trimmed = contents.trim();
      if (trimmed.length === 0) {
        setStickyDraft(null);
        onAnnotationCreated?.();
        return;
      }
      const { x, y } = stickyDraft;
      const half = 9; // ~18 PDF units square pin, like Acrobat
      const r1 = Math.max(0, x - half);
      const r2 = Math.min(page.width, x + half);
      const r3 = Math.max(0, y - half);
      const r4 = Math.min(page.height, y + half);
      onAddAnnotation?.({
        kind: "note",
        pageNumber: page.pageNumber,
        rect: [r1, r3, r2, r4] as PdfRect,
        contents: trimmed,
        color: { r: 1, g: 0.84, b: 0.0 },
      });
      setStickyDraft(null);
      onAnnotationCreated?.();
    },
    [stickyDraft, onAddAnnotation, onAnnotationCreated, page.pageNumber, page.width, page.height]
  );
  // Clearing the draft when un-arming keeps the composer from
  // surviving an Escape press through the editor's keyboard wiring.
  useEffect(() => {
    if (!stickyArmed) setStickyDraft(null);
  }, [stickyArmed]);
  const { t } = useTranslator();
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const renderTokenRef = useRef(0);
  const [textItems, setTextItems] = useState<ReadonlyArray<PdfEngineTextItem>>([]);
  const [renderError, setRenderError] = useState<string | null>(null);

  const cssWidth = Math.max(1, Math.round(effectiveWidth(page, viewportRotation) * scale));
  const cssHeight = Math.max(1, Math.round(effectiveHeight(page, viewportRotation) * scale));

  useEffect(() => {
    if (!visible) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;
    const dpr = typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1;
    const token = ++renderTokenRef.current;
    let cancelled = false;
    let enginePage: PdfEnginePage | null = null;
    setRenderError(null);

    void (async () => {
      try {
        enginePage = await engineDoc.getPage(page.pageNumber);
        if (cancelled || token !== renderTokenRef.current) return;
        canvas.width = Math.max(1, Math.floor(cssWidth * dpr));
        canvas.height = Math.max(1, Math.floor(cssHeight * dpr));
        canvas.style.width = `${cssWidth}px`;
        canvas.style.height = `${cssHeight}px`;
        ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
        ctx.fillStyle = "white";
        ctx.fillRect(0, 0, cssWidth, cssHeight);
        await enginePage.render({
          canvas,
          scale: scale * dpr,
          rotation: viewportRotation,
        });
        if (cancelled || token !== renderTokenRef.current) return;

        try {
          const text = await enginePage.getTextContent();
          if (cancelled || token !== renderTokenRef.current) return;
          setTextItems(text.items);
        } catch {
          // Text extraction failures are non-fatal — selectable text
          // simply won't appear for this page. The toast is reserved
          // for whole-page render failures below.
        }
      } catch (err) {
        if (cancelled) return;
        setRenderError(err instanceof Error ? err.message : String(err));
      } finally {
        try {
          enginePage?.destroy();
        } catch {
          /* noop */
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [visible, engineDoc, page.pageNumber, scale, cssWidth, cssHeight, darkMode, viewportRotation]);

  const cursorClass = highlightArmed
    ? "cursor-text"
    : stickyArmed
    ? "cursor-crosshair"
    : "cursor-default";

  return (
    <div
      data-page-number={page.pageNumber}
      data-testid={`pdf-page-${page.pageNumber}`}
      onClick={onClick}
      className={
        "relative rounded-md border bg-white shadow-sm " +
        cursorClass +
        " " +
        (active ? "border-[var(--accent)]" : "border-divider")
      }
      style={{ width: cssWidth, height: cssHeight }}
    >
      <canvas
        ref={canvasRef}
        className="block h-full w-full rounded-md"
        style={{
          filter: darkMode === "on" ? darkModeCssFilter(true) : "none",
        }}
        aria-label={t("pdf.pageOf", { n: page.pageNumber, total: totalPages })}
      />
      {visible && textItems.length > 0 ? (
        <PdfTextLayer
          items={textItems}
          scale={scale}
          pageWidth={page.width}
          pageHeight={page.height}
          rotation={viewportRotation}
          onSelectionPdfRects={highlightArmed ? onHighlightSelection : undefined}
        />
      ) : null}
      {visible && annotations.length > 0 ? (
        <PdfAnnotationOverlay
          annotations={annotations}
          pageWidth={page.width}
          pageHeight={page.height}
          scale={scale}
          rotation={viewportRotation}
        />
      ) : null}
      {visible ? (
        <PdfStickyNotesOverlay
          annotations={annotations}
          pageWidth={page.width}
          pageHeight={page.height}
          scale={scale}
          rotation={viewportRotation}
        />
      ) : null}
      {stickyArmed ? (
        <div
          aria-hidden
          className="absolute inset-0 z-10"
          onClick={onPageClickForSticky}
        />
      ) : null}
      {stickyDraft && viewportRotation === 0 ? (
        <PdfStickyComposer
          x={stickyDraft.x * scale}
          y={(page.height - stickyDraft.y) * scale}
          onCommit={onCommitSticky}
          onCancel={() => setStickyDraft(null)}
        />
      ) : null}
      {highlight ? <PdfMatchHighlight highlight={highlight} /> : null}
      {renderError ? (
        <div className="absolute inset-0 flex items-center justify-center rounded-md bg-surface/90 p-3 text-center text-xs text-tertiary">
          {t("pdf.loadError")}: {renderError}
        </div>
      ) : null}
      {!page.hasTextLayer && visible ? (
        <div
          className="pointer-events-none absolute bottom-1 right-1 rounded bg-surface/80 px-1.5 py-0.5 text-[10px] text-tertiary"
          aria-label={t("pdf.scanned")}
        >
          {t("pdf.scanned")}
        </div>
      ) : null}
    </div>
  );
}

interface PdfTextLayerProps {
  readonly items: ReadonlyArray<PdfEngineTextItem>;
  readonly scale: number;
  readonly pageWidth: number;
  readonly pageHeight: number;
  readonly rotation: PdfRotation;
  /**
   * If provided, mouseup events whose selection ends inside this
   * text layer fire this callback with the selection's per-line
   * rects translated into PDF user-space (origin bottom-left,
   * units = PDF points). Only emitted at rotation 0 — the editor
   * gates the highlight tool on that.
   */
  readonly onSelectionPdfRects?: (rects: ReadonlyArray<PdfRect>) => void;
}

/**
 * Selectable, transparent text overlay on top of the rasterised
 * page. We synthesise one absolutely-positioned `<span>` per
 * `PdfEngineTextItem` placed at the item's PDF user-space
 * coordinates, scaled into CSS pixels and y-flipped to the DOM's
 * top-left origin. The font size is taken from the item transform
 * so the synthesised glyphs roughly cover the rendered ones — that
 * makes click-and-drag selection feel native even though the
 * actual glyph shapes are rasterised by the engine.
 *
 * The inner box is sized to the un-rotated page dimensions and
 * then CSS-rotated to match the viewport rotation that the engine
 * baked into the canvas, so spans always land on top of the
 * corresponding glyphs.
 */
function PdfTextLayer({
  items,
  scale,
  pageWidth,
  pageHeight,
  rotation,
  onSelectionPdfRects,
}: PdfTextLayerProps): ReactNode {
  const innerWidth = pageWidth * scale;
  const innerHeight = pageHeight * scale;
  const rotatedW = rotation === 90 || rotation === 270 ? innerHeight : innerWidth;
  const rotatedH = rotation === 90 || rotation === 270 ? innerWidth : innerHeight;
  const transform = textLayerTransform(rotation, innerWidth, innerHeight);
  const innerRef = useRef<HTMLDivElement | null>(null);

  const onMouseUp = useCallback(() => {
    if (!onSelectionPdfRects) return;
    const inner = innerRef.current;
    if (!inner) return;
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    if (!sel || sel.isCollapsed || sel.rangeCount === 0) return;
    const range = sel.getRangeAt(0);
    if (!inner.contains(range.commonAncestorContainer)) return;
    const innerRect = inner.getBoundingClientRect();
    if (innerRect.width === 0 || innerRect.height === 0) return;
    const clientRects = Array.from(range.getClientRects()).filter(
      (r) => r.width > 0 && r.height > 0
    );
    if (clientRects.length === 0) return;
    const merged = mergeRectsByLine(clientRects);
    const pdfRects: PdfRect[] = merged.map((r) => {
      const x_css = r.left - innerRect.left;
      const y_css = r.top - innerRect.top;
      const w_css = r.width;
      const h_css = r.height;
      const x1 = x_css / scale;
      const x2 = (x_css + w_css) / scale;
      const y2 = pageHeight - y_css / scale;
      const y1 = pageHeight - (y_css + h_css) / scale;
      return [x1, y1, x2, y2] as PdfRect;
    });
    sel.removeAllRanges();
    onSelectionPdfRects(pdfRects);
  }, [onSelectionPdfRects, scale, pageHeight]);

  // PDF.js's reference text layer measures each synthesised span's
  // browser-rendered width against the desired (rasterised) glyph
  // width and applies a per-span `transform: scaleX(target/measured)`
  // so the invisible selection rect lines up with the visible
  // letters. We do the same: without it, click-drag selection
  // floats noticeably to the right of the rasterised text on most
  // body fonts because the platform's default sans-serif is wider
  // than the PDF's embedded font.
  useLayoutEffect(() => {
    const root = innerRef.current;
    if (!root) return;
    const spans = root.querySelectorAll<HTMLSpanElement>("[data-text-span]");
    spans.forEach((span) => {
      span.style.transform = "";
      const target = Number(span.dataset.targetWidth ?? "0");
      if (!target) return;
      const measured = span.getBoundingClientRect().width;
      if (measured <= 0) return;
      const ratio = target / measured;
      // Only correct when the discrepancy matters; tiny ratios just
      // burn paint cycles for sub-pixel adjustments invisible to
      // the eye but expensive on long documents.
      if (Math.abs(ratio - 1) < 0.01) return;
      span.style.transform = `scaleX(${ratio.toFixed(4)})`;
    });
  }, [items, scale, rotation]);

  return (
    <div
      aria-hidden
      onMouseUp={onMouseUp}
      className="pointer-events-auto absolute inset-0 select-text overflow-hidden text-transparent"
      style={{ lineHeight: 1, width: rotatedW, height: rotatedH }}
    >
      <div
        ref={innerRef}
        style={{
          position: "absolute",
          left: 0,
          top: 0,
          width: innerWidth,
          height: innerHeight,
          transform,
          transformOrigin: "0 0",
        }}
      >
        {items.map((it, idx) => {
          if (!it.str) return null;
          const tx = it.transform;
          const fontHeight = Math.abs(tx[3]) || it.height || 1;
          const left = tx[4] * scale;
          const top = (pageHeight - tx[5] - fontHeight) * scale;
          const targetWidth = it.width * scale;
          const height = fontHeight * scale;
          return (
            <span
              key={idx}
              data-text-span
              data-target-width={targetWidth}
              style={{
                position: "absolute",
                left,
                top,
                height,
                fontSize: height,
                whiteSpace: "pre",
                transformOrigin: "0 0",
              }}
            >
              {it.str}
            </span>
          );
        })}
      </div>
    </div>
  );
}

/**
 * CSS transform that maps the un-rotated text-layer box onto the
 * rotated visual frame. Mirrors PDF.js's viewport rotation so
 * synthesised spans line up with the rasterised glyphs.
 */
function textLayerTransform(rotation: PdfRotation, width: number, height: number): string {
  switch (rotation) {
    case 0:
      return "none";
    case 90:
      return `translate(${height}px, 0) rotate(90deg)`;
    case 180:
      return `translate(${width}px, ${height}px) rotate(180deg)`;
    case 270:
      return `translate(0, ${width}px) rotate(270deg)`;
    default: {
      const _exhaustive: never = rotation;
      return _exhaustive;
    }
  }
}

interface PdfAnnotationOverlayProps {
  readonly annotations: ReadonlyArray<PdfAnnotation>;
  readonly pageWidth: number;
  readonly pageHeight: number;
  readonly scale: number;
  readonly rotation: PdfRotation;
}

/**
 * Renders translucent overlays for highlight annotations on the
 * current page. Sticky-note pins and link affordances are layered
 * separately by their own components in WP9; for WP8 we only need
 * to surface highlights so the user immediately sees the result of
 * arming the highlight tool and dragging across text.
 *
 * Coordinates: annotations live in PDF user-space (origin
 * bottom-left). We position each rect inside an inner box of the
 * un-rotated page dimensions and let the same CSS transform that
 * rotates the text layer rotate the overlay too — so they never
 * drift out of sync.
 */
function PdfAnnotationOverlay({
  annotations,
  pageWidth,
  pageHeight,
  scale,
  rotation,
}: PdfAnnotationOverlayProps): ReactNode {
  const innerWidth = pageWidth * scale;
  const innerHeight = pageHeight * scale;
  const rotatedW = rotation === 90 || rotation === 270 ? innerHeight : innerWidth;
  const rotatedH = rotation === 90 || rotation === 270 ? innerWidth : innerHeight;
  const transform = textLayerTransform(rotation, innerWidth, innerHeight);
  const highlights = annotations.filter((a) => a.kind === "highlight");
  if (highlights.length === 0) return null;
  return (
    <div
      aria-hidden
      className="pointer-events-none absolute inset-0"
      style={{ width: rotatedW, height: rotatedH }}
    >
      <div
        style={{
          position: "absolute",
          left: 0,
          top: 0,
          width: innerWidth,
          height: innerHeight,
          transform,
          transformOrigin: "0 0",
        }}
      >
        {highlights.map((a) => {
          const rects = a.quadRects && a.quadRects.length > 0 ? a.quadRects : [a.rect];
          return rects.map((r, i) => {
            const [x1, y1, x2, y2] = r;
            const left = x1 * scale;
            const top = (pageHeight - y2) * scale;
            const width = Math.max(1, (x2 - x1) * scale);
            const height = Math.max(1, (y2 - y1) * scale);
            const fill = annotationFill(a);
            return (
              <div
                key={`${a.id}-${i}`}
                style={{
                  position: "absolute",
                  left,
                  top,
                  width,
                  height,
                  background: fill,
                  mixBlendMode: "multiply",
                }}
              />
            );
          });
        })}
      </div>
    </div>
  );
}

function annotationFill(a: PdfAnnotation): string {
  const c = a.color;
  if (!c) return "rgba(255, 235, 59, 0.45)";
  const alpha = c.a ?? 0.45;
  const r = Math.round((c.r ?? 1) * 255);
  const g = Math.round((c.g ?? 0.93) * 255);
  const b = Math.round((c.b ?? 0.16) * 255);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function pageAnnotations(
  snapshot: PdfSnapshot,
  pageNumber: number
): ReadonlyArray<PdfAnnotation> {
  return snapshot.root.annotations.filter((a) => a.pageNumber === pageNumber);
}

/** Axis-aligned union of a set of PDF rects. */
function unionPdfRects(rects: ReadonlyArray<PdfRect>): PdfRect {
  let x1 = Number.POSITIVE_INFINITY;
  let y1 = Number.POSITIVE_INFINITY;
  let x2 = Number.NEGATIVE_INFINITY;
  let y2 = Number.NEGATIVE_INFINITY;
  for (const r of rects) {
    if (r[0] < x1) x1 = r[0];
    if (r[1] < y1) y1 = r[1];
    if (r[2] > x2) x2 = r[2];
    if (r[3] > y2) y2 = r[3];
  }
  return [x1, y1, x2, y2] as PdfRect;
}

interface MergedRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/**
 * `Range.getClientRects()` returns one rect per visual line — but
 * sometimes splits a single line into many adjacent rects when the
 * selection traverses span boundaries. Merging them by row keeps
 * the highlight quad count low and avoids visible seams between
 * sub-rects on the same baseline.
 */
function mergeRectsByLine(rects: ReadonlyArray<DOMRect>): ReadonlyArray<MergedRect> {
  if (rects.length === 0) return [];
  const sorted = [...rects].sort((a, b) => a.top - b.top || a.left - b.left);
  const merged: MergedRect[] = [];
  for (const r of sorted) {
    const last = merged[merged.length - 1];
    if (last && Math.abs(r.top - last.top) < r.height * 0.5) {
      const newLeft = Math.min(last.left, r.left);
      const newRight = Math.max(last.left + last.width, r.left + r.width);
      const newTop = Math.min(last.top, r.top);
      const newBottom = Math.max(last.top + last.height, r.top + r.height);
      last.left = newLeft;
      last.top = newTop;
      last.width = newRight - newLeft;
      last.height = newBottom - newTop;
    } else {
      merged.push({ left: r.left, top: r.top, width: r.width, height: r.height });
    }
  }
  return merged;
}

interface PdfStickyNotesOverlayProps {
  readonly annotations: ReadonlyArray<PdfAnnotation>;
  readonly pageWidth: number;
  readonly pageHeight: number;
  readonly scale: number;
  readonly rotation: PdfRotation;
}

/**
 * Renders sticky-note pins for any `note` annotations on the page.
 * Each pin opens an inline popover on click that displays its
 * contents. Editing/deleting lives behind the same popover so
 * users never need to leave the page rail.
 *
 * The pin layer ignores rotation overlay logic deliberately: PDF
 * note rects use bottom-left origin, and we only commit pins from
 * rotation 0. Existing pins stay anchored to their PDF coords —
 * if the user rotates the viewport, pins ride along via the same
 * inner-box CSS transform.
 */
function PdfStickyNotesOverlay({
  annotations,
  pageWidth,
  pageHeight,
  scale,
  rotation,
}: PdfStickyNotesOverlayProps): ReactNode {
  const innerWidth = pageWidth * scale;
  const innerHeight = pageHeight * scale;
  const rotatedW = rotation === 90 || rotation === 270 ? innerHeight : innerWidth;
  const rotatedH = rotation === 90 || rotation === 270 ? innerWidth : innerHeight;
  const transform = textLayerTransform(rotation, innerWidth, innerHeight);
  const [openId, setOpenId] = useState<string | null>(null);
  const notes = annotations.filter((a) => a.kind === "note");
  if (notes.length === 0) return null;
  return (
    <div
      className="pointer-events-none absolute inset-0"
      style={{ width: rotatedW, height: rotatedH }}
    >
      <div
        style={{
          position: "absolute",
          left: 0,
          top: 0,
          width: innerWidth,
          height: innerHeight,
          transform,
          transformOrigin: "0 0",
        }}
      >
        {notes.map((a) => {
          const [x1, , x2, y2] = a.rect;
          const cx = ((x1 + x2) / 2) * scale;
          const cy = (pageHeight - y2) * scale;
          const open = openId === a.id;
          return (
            <div key={a.id} style={{ position: "absolute", left: cx - 10, top: cy - 10 }}>
              <button
                type="button"
                title={a.contents ?? ""}
                aria-label={a.contents ?? "sticky note"}
                onClick={(e) => {
                  e.stopPropagation();
                  setOpenId((current) => (current === a.id ? null : a.id));
                }}
                className="pointer-events-auto flex h-5 w-5 items-center justify-center rounded-sm border border-amber-600/60 bg-amber-300 text-[10px] text-amber-900 shadow-sm hover:bg-amber-200"
                style={{ pointerEvents: "auto" }}
              >
                ✎
              </button>
              {open ? (
                <div
                  className="pointer-events-auto absolute left-6 top-0 z-20 w-56 rounded-md border border-amber-300 bg-amber-50 p-2 text-xs text-amber-950 shadow-lg"
                  onClick={(e) => e.stopPropagation()}
                >
                  <div className="whitespace-pre-wrap break-words">
                    {a.contents || "(no contents)"}
                  </div>
                  {a.author ? (
                    <div className="mt-1 text-[10px] uppercase tracking-wide text-amber-700/80">
                      {a.author}
                    </div>
                  ) : null}
                </div>
              ) : null}
            </div>
          );
        })}
      </div>
    </div>
  );
}

interface PdfStickyComposerProps {
  /** CSS x of the click in page-local coordinates. */
  readonly x: number;
  /** CSS y of the click in page-local coordinates. */
  readonly y: number;
  readonly onCommit: (contents: string) => void;
  readonly onCancel: () => void;
}

/**
 * Inline composer for a brand-new sticky note. Auto-focuses on
 * mount; Cmd/Ctrl+Enter commits, Esc cancels. Positioned via
 * absolute coordinates relative to the page wrapper.
 */
function PdfStickyComposer({ x, y, onCommit, onCancel }: PdfStickyComposerProps): ReactNode {
  const { t } = useTranslator();
  const [value, setValue] = useState("");
  const ref = useRef<HTMLTextAreaElement | null>(null);
  useEffect(() => {
    ref.current?.focus();
  }, []);
  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>): void => {
    if (e.key === "Escape") {
      e.stopPropagation();
      onCancel();
    } else if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) {
      e.preventDefault();
      onCommit(value);
    }
  };
  return (
    <div
      className="absolute z-30"
      style={{ left: x, top: y }}
      onClick={(e) => e.stopPropagation()}
    >
      <div className="w-56 rounded-md border border-amber-300 bg-amber-50 p-2 shadow-xl">
        <textarea
          ref={ref}
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={onKeyDown}
          rows={4}
          placeholder={t("pdf.stickyPlaceholder")}
          className="w-full resize-none rounded border border-amber-300/80 bg-white p-1.5 text-xs text-amber-950 outline-none focus:border-amber-500"
        />
        <div className="mt-1.5 flex items-center justify-between">
          <span className="text-[10px] text-amber-700/80">⌘↵</span>
          <div className="flex gap-1">
            <button
              type="button"
              onClick={onCancel}
              className="rounded border border-transparent px-2 py-0.5 text-[11px] text-amber-900 hover:bg-amber-100"
            >
              {t("pdf.stickyCancel")}
            </button>
            <button
              type="button"
              onClick={() => onCommit(value)}
              className="rounded bg-amber-500 px-2 py-0.5 text-[11px] font-medium text-white hover:bg-amber-600"
            >
              {t("pdf.stickySave")}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function PdfMatchHighlight({ highlight }: { readonly highlight: PdfHighlight }): ReactNode {
  const [x1, y1, x2, y2] = highlight.rect;
  const left = `${x1 * 100}%`;
  const top = `${y1 * 100}%`;
  const width = `${(x2 - x1) * 100}%`;
  const height = `${(y2 - y1) * 100}%`;
  return (
    <div
      key={highlight.nonce}
      className="pointer-events-none absolute rounded-sm bg-[var(--accent-light)] ring-2 ring-[var(--accent)]/60"
      style={{ left, top, width, height }}
    />
  );
}

interface ReflowPageProps {
  readonly page: PdfPage;
  readonly totalPages: number;
}

function ReflowPage({ page, totalPages }: ReflowPageProps): ReactNode {
  const { t } = useTranslator();
  const heading = t("pdf.pageOf", { n: page.pageNumber, total: totalPages });
  if (!page.text || !page.hasTextLayer) {
    return (
      <section className="my-6 rounded-md border border-divider bg-surface px-4 py-3 text-xs text-tertiary">
        {t("pdf.scanned")} — {heading}
      </section>
    );
  }
  return (
    <section className="my-6">
      <h2 className="mb-2 text-xs font-semibold uppercase tracking-wider text-tertiary">{heading}</h2>
      <pre className="whitespace-pre-wrap font-sans text-sm text-primary">{page.text}</pre>
    </section>
  );
}

/** Compute the rendered width of a page after the user's viewport
 * rotation. Intrinsic page rotation is already baked into the
 * engine's render output; we only swap axes for *additional* 90° /
 * 270° viewport rotation. */
function effectiveWidth(page: PdfPage, viewportRotation: PdfRotation): number {
  return viewportRotation === 90 || viewportRotation === 270 ? page.height : page.width;
}
function effectiveHeight(page: PdfPage, viewportRotation: PdfRotation): number {
  return viewportRotation === 90 || viewportRotation === 270 ? page.width : page.height;
}

interface RenderRow {
  readonly key: string;
  readonly kind: "single" | "two-up";
  readonly pages: ReadonlyArray<PdfPage>;
}

function collectRenderTargets(
  pages: ReadonlyArray<PdfPage>,
  viewMode: PdfViewMode,
  currentPage: number
): ReadonlyArray<RenderRow> {
  if (viewMode === "single") {
    const p = pages.find((x) => x.pageNumber === currentPage) ?? pages[0];
    return p ? [{ key: `single-${p.id}`, kind: "single", pages: [p] }] : [];
  }
  if (viewMode === "two-up") {
    const rows: RenderRow[] = [];
    if (pages[0]) {
      rows.push({ key: `cover-${pages[0].id}`, kind: "two-up", pages: [pages[0]] });
    }
    for (let i = 1; i < pages.length; i += 2) {
      const left = pages[i];
      if (!left) break;
      const right = pages[i + 1];
      const inRow = right ? [left, right] : [left];
      rows.push({ key: `pair-${left.id}`, kind: "two-up", pages: inRow });
    }
    return rows;
  }
  return pages.map((p) => ({ key: `cont-${p.id}`, kind: "single", pages: [p] }));
}
