"use client";

import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import type { PdfEngineDocument, PdfEnginePage } from "@officeai/pdf-engine";
import type {
  AddAnnotationPayload,
  PdfAnnotation,
  PdfPage,
  PdfRect,
  PdfRotation,
  PdfSnapshot,
} from "@officeai/pdf";
import { collectTextWithinRegions } from "@officeai/pdf";
import { useTranslator } from "@/lib/i18n";
import { darkModeCssFilter } from "./darkMode";
import type { PdfAnnotationTool } from "./PdfToolbar";
import "./textLayer.css";

/** Canvas-side view modes — 1:1 with the toolbar enum. */
export type PdfViewMode = "single" | "continuous" | "two-up";

/**
 * Gap between the two facing pages of a two-up spread, in CSS
 * pixels. Kept tiny so the spread reads as an open book with a
 * binding seam rather than two unrelated pages floating apart.
 * The value is also subtracted from the per-page fit budget in
 * `baseScale` so the spread always fits the container.
 */
const BOOK_BINDING_GAP = 2;

/**
 * Binary dark-mode toggle. CSS-filter based, no per-pixel pass.
 */
export type PdfDarkModeStrategy = "off" | "on";

export interface PdfHighlight {
  readonly pageNumber: number;
  /** Normalized fallback rect [x1, y1, x2, y2] (0..1, origin top-left).
   * Used when `quads` are unavailable (regex hits across paragraphs,
   * scanned pages, etc.) — the canvas paints a soft full-width pulse. */
  readonly rect: readonly [number, number, number, number];
  /**
   * Per-line bounding boxes of the matched glyphs in PDF user-space
   * (origin bottom-left). When set, the canvas paints these over the
   * actual text instead of the page-level pulse — i.e. real
   * find-in-page highlights instead of "we found something on this
   * page somewhere".
   */
  readonly quads?: ReadonlyArray<PdfRect>;
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
  const widestPageDim = useMemo(() => {
    let widest = 0;
    for (const p of pages) {
      const w = effectiveWidth(p, viewportRotation);
      if (w > widest) widest = w;
    }
    return widest;
  }, [pages, viewportRotation]);
  const baseScale = useMemo(() => {
    if (!pages.length || containerWidth === 0 || widestPageDim === 0) return 1;
    // Fit the widest page (or a side-by-side pair, in two-up mode)
    // into the container with a small gutter on each side. The
    // user's `zoom` then multiplies on top.
    //
    // In two-up the row places two pages touching at the binding
    // (separated by `BOOK_BINDING_GAP`), so the available width
    // *per page* is (container − gutter − gap) / 2. Without
    // dividing by 2 here a pair would render at roughly 2× the
    // container width, overflow horizontally, and (combined with
    // flex `min-width: auto` on the ancestor wrappers) feed back
    // into ResizeObserver — `containerWidth` would explode to the
    // browser's layout cap and the `<canvas>` would die as a
    // broken image icon.
    const gutter = viewMode === "two-up" ? 24 : 48;
    const target = Math.max(120, containerWidth - gutter);
    if (viewMode === "two-up") {
      const perPage = Math.max(60, (target - BOOK_BINDING_GAP) / 2);
      return perPage / widestPageDim;
    }
    return target / widestPageDim;
  }, [pages.length, widestPageDim, containerWidth, viewMode]);
  const scale = baseScale * zoom;
  // The slot width every two-up page (and every spacer) takes up.
  // Pinning all rows to this value keeps the spread visually
  // aligned even when individual page sizes differ across the
  // document — narrower pages just inherit the slot, the canvas
  // itself is still drawn at its native ratio.
  const twoUpSlotWidth = Math.max(1, Math.round(widestPageDim * scale));

  // Surface fit-page / actual-size targets back to the editor so
  // the toolbar's preset callbacks set a real `zoom` instead of a
  // hard-coded heuristic.
  useEffect(() => {
    if (!onZoomMetricsChange) return;
    if (!pages.length || baseScale === 0) return;
    // Wait until the container has been measured in *both* axes —
    // emitting fit-page metrics with `containerHeight === 0` would
    // give the editor a bogus fitPageZoom of 1 and (since the
    // editor consumes only the first metrics callback after a
    // document load) lock the page in at "way too zoomed in".
    if (containerHeight === 0) return;
    const tallest = pages.reduce((acc, p) => Math.max(acc, effectiveHeight(p, viewportRotation)), 0);
    if (tallest === 0) return;
    const verticalGutter = 32;
    const fitPageScale = Math.min(baseScale, Math.max(0.05, (containerHeight - verticalGutter) / tallest));
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
        const placeholders = Array.from(root.querySelectorAll<HTMLElement>("[data-page-number]"));
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
    const target = root.querySelector<HTMLElement>(`[data-page-number="${currentPage}"]`);
    if (!target) return;
    const top = target.offsetTop;
    const distance = Math.abs(top - root.scrollTop);
    if (distance < 4) return;
    // Smooth feels great for small jumps; for big jumps (e.g.
    // shift-jumping 50 pages) the animation drags. Snap instantly
    // past five page-heights.
    const tallest = pages.reduce((acc, p) => Math.max(acc, effectiveHeight(p, viewportRotation) * scale), 1);
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
  // Fully-scanned ⇒ no page reports a text layer. We deliberately
  // don't fire the banner for partially-scanned documents (e.g.
  // a TOC + scanned chapters) — pages that *do* have text behave
  // normally and per-page badges would just add noise.
  const isFullyScanned = pages.length > 0 && pages.every((p) => !p.hasTextLayer);

  return (
    // `min-w-0` defeats the default `min-width: auto` on flex
    // children: without it, a row that briefly renders wider than
    // the container (e.g. mid-resize, before `baseScale` settles)
    // would push the flex parent — and therefore `clientWidth` —
    // out, feeding back into `baseScale` and snowballing until the
    // browser's 2²⁴-px layout cap clamps the canvas to a broken
    // image. `overflow-x-hidden` belt-and-braces the same guard
    // for non-flex ancestors.
    <div
      ref={containerRef}
      className="h-full w-full min-w-0 overflow-y-auto overflow-x-hidden bg-[var(--page-backdrop)]"
      data-testid="pdf-canvas"
    >
      {isFullyScanned ? (
        <div
          className="mx-auto mt-4 flex w-full max-w-3xl items-start gap-2 rounded-md border border-amber-300 bg-amber-50 px-3 py-2 text-xs text-amber-950"
          data-testid="pdf-scanned-banner"
        >
          <span className="mt-0.5 font-semibold">{t("pdf.scannedBannerTitle")}</span>
          <span className="text-amber-900/90">{t("pdf.scannedBannerHint")}</span>
        </div>
      ) : null}
      <div className="mx-auto flex max-w-full flex-col items-center gap-4 py-6">
        {renderTargets.map((row) => {
          const renderPage = (p: PdfPage): ReactNode => (
            <PdfPageRender
              key={p.id}
              page={p}
              totalPages={totalPages}
              engineDoc={engineDoc}
              scale={scale}
              visible={visibleRef.current.has(p.pageNumber)}
              viewportRotation={viewportRotation}
              darkMode={darkMode}
              onClick={() => onCurrentPageChange(p.pageNumber)}
              highlight={highlight && highlight.pageNumber === p.pageNumber ? highlight : null}
              annotations={pageAnnotations(snapshot, p.pageNumber)}
              armedTool={armedTool}
              onAddAnnotation={onAddAnnotation}
              onAnnotationCreated={onAnnotationCreated}
            />
          );
          if (row.kind !== "two-up") {
            return <div key={row.key}>{row.pages.map(renderPage)}</div>;
          }
          // Two-up rows render at a fixed spread width regardless of
          // how many pages are in the slot, so the cover, the
          // trailing lone page, and the full spreads all share the
          // same horizontal axis. Empty slots are filled with a
          // transparent placeholder of the same dimensions; that
          // also gives the cover its "right side of an open book"
          // offset for free.
          const spreadWidth = twoUpSlotWidth * 2 + BOOK_BINDING_GAP;
          const leftPage = row.slot === "right" ? null : (row.pages[0] ?? null);
          const rightPage =
            row.slot === "spread"
              ? (row.pages[1] ?? null)
              : row.slot === "right"
                ? (row.pages[0] ?? null)
                : null;
          return (
            <div
              key={row.key}
              data-testid={`pdf-spread-${row.slot}`}
              className="flex items-start"
              style={{ width: spreadWidth, gap: BOOK_BINDING_GAP }}
            >
              {leftPage ? (
                renderPage(leftPage)
              ) : (
                <div aria-hidden style={{ width: twoUpSlotWidth, flexShrink: 0 }} />
              )}
              {rightPage ? (
                renderPage(rightPage)
              ) : (
                <div aria-hidden style={{ width: twoUpSlotWidth, flexShrink: 0 }} />
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

interface PdfPageRenderProps {
  readonly page: PdfPage;
  readonly totalPages: number;
  readonly engineDoc: PdfEngineDocument;
  readonly scale: number;
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
  const textLayerRef = useRef<HTMLDivElement | null>(null);
  const renderTokenRef = useRef(0);
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
    let textLayerHandle: TextLayerHandle | null = null;
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

        const layerEl = textLayerRef.current;
        if (layerEl) {
          try {
            textLayerHandle = await mountOfficialTextLayer({
              enginePage,
              container: layerEl,
              scale,
              rotation: viewportRotation,
            });
            if (cancelled || token !== renderTokenRef.current) {
              textLayerHandle?.cancel();
              textLayerHandle = null;
            }
          } catch {
            // Text-layer failures are non-fatal — selectable text
            // simply won't appear for this page. The toast is
            // reserved for whole-page render failures below.
          }
        }
      } catch (err) {
        if (cancelled) return;
        setRenderError(err instanceof Error ? err.message : String(err));
      } finally {
        // Hold engine page open for the layer's lifetime, but it's
        // safe to release here because PDF.js's TextLayer keeps its
        // own references to the streaming text content.
        try {
          enginePage?.destroy();
        } catch {
          /* noop */
        }
      }
    })();

    return () => {
      cancelled = true;
      textLayerHandle?.cancel();
      textLayerHandle = null;
    };
  }, [visible, engineDoc, page.pageNumber, scale, cssWidth, cssHeight, darkMode, viewportRotation]);

  const cursorClass = highlightArmed ? "cursor-text" : stickyArmed ? "cursor-crosshair" : "cursor-default";

  return (
    <div
      data-page-number={page.pageNumber}
      data-testid={`pdf-page-${page.pageNumber}`}
      onClick={onClick}
      className={"relative rounded-md border border-divider bg-white shadow-sm " + cursorClass}
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
      <OfficialPdfTextLayer
        ref={textLayerRef}
        scale={scale}
        pageWidth={page.width}
        pageHeight={page.height}
        rotation={viewportRotation}
        page={page}
        onSelectionPdfRects={highlightArmed ? onHighlightSelection : undefined}
      />

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
        <div aria-hidden className="absolute inset-0 z-10" onClick={onPageClickForSticky} />
      ) : null}
      {stickyDraft && viewportRotation === 0 ? (
        <PdfStickyComposer
          x={stickyDraft.x * scale}
          y={(page.height - stickyDraft.y) * scale}
          onCommit={onCommitSticky}
          onCancel={() => setStickyDraft(null)}
        />
      ) : null}
      {highlight ? (
        <PdfMatchHighlight
          highlight={highlight}
          pageWidth={page.width}
          pageHeight={page.height}
          scale={scale}
          rotation={viewportRotation}
        />
      ) : null}
      {renderError ? (
        <div className="absolute inset-0 flex items-center justify-center rounded-md bg-surface/90 p-3 text-center text-xs text-tertiary">
          {t("pdf.loadError")}: {renderError}
        </div>
      ) : null}
      {/* The per-page "scanned" badge was removed in WP10 — it
          fired far too eagerly (any page without a text layer,
          including blank cover pages). The document-level banner
          surfaces the same hint exactly when it's actionable. */}
    </div>
  );
}

interface OfficialPdfTextLayerProps {
  readonly scale: number;
  readonly pageWidth: number;
  readonly pageHeight: number;
  readonly rotation: PdfRotation;
  /** The owning page — needed to project a selection into the
   * structured text for column-aware copy interception. */
  readonly page: PdfPage;
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
 * Selectable, transparent text overlay backed by PDF.js's official
 * `TextLayer` class. The container `<div>` is the same DOM node
 * PDF.js writes its `<span>`s into; positioning, font matching,
 * baseline ascent and per-glyph width compensation are all handled
 * by the canonical layer instead of the bespoke approximation we
 * used to ship.
 *
 * The CSS variable `--scale-factor` is what PDF.js's TextLayer
 * uses for its internal CSS calc()s — without it the spans render
 * at fixed (1.0) scale and drift the moment the user zooms.
 */
const OfficialPdfTextLayer = React.forwardRef<HTMLDivElement, OfficialPdfTextLayerProps>(
  function OfficialPdfTextLayer({ scale, pageWidth, pageHeight, rotation, page, onSelectionPdfRects }, ref) {
    const innerWidth = pageWidth * scale;
    const innerHeight = pageHeight * scale;
    const rotatedW = rotation === 90 || rotation === 270 ? innerHeight : innerWidth;
    const rotatedH = rotation === 90 || rotation === 270 ? innerWidth : innerHeight;

    /**
     * Reorder copied text along the structured-page reading order
     * before letting the clipboard event ship. Without this, selecting
     * across columns on academic PDFs pastes interleaved garbage
     * because the browser concatenates `<span>`s in DOM order — which
     * is also stream order, i.e. left-column-line-1, right-column-
     * line-1, left-column-line-2, ... For single-column pages and
     * intra-paragraph selections we leave the browser's text alone
     * (it's already correct and faster).
     */
    const onCopy = useCallback(
      (e: React.ClipboardEvent<HTMLDivElement>) => {
        const inner = typeof ref === "object" && ref ? ref.current : null;
        if (!inner) return;
        const sel = typeof window !== "undefined" ? window.getSelection() : null;
        if (!sel || sel.isCollapsed || sel.rangeCount === 0) return;
        const range = sel.getRangeAt(0);
        if (!inner.contains(range.commonAncestorContainer)) return;
        if (page.structured.blocks.length === 0) return;

        const innerRect = inner.getBoundingClientRect();
        if (innerRect.width === 0 || innerRect.height === 0) return;
        const clientRects = Array.from(range.getClientRects()).filter((r) => r.width > 0 && r.height > 0);
        if (clientRects.length === 0) return;
        // Convert the selection's pixel rects into PDF user-space
        // and ask the structured page which glyphs land inside.
        const pdfRegions = clientRects.map((r) => {
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
        const text = collectTextWithinRegions(page.structured, pdfRegions);
        const browserText = sel.toString();
        // Only override when the structured projection meaningfully
        // differs (multi-column selection, or row-wise jumble).
        // For single-column intra-paragraph selections both texts
        // agree, in which case we let the browser do its thing for
        // free.
        if (text.length === 0) return;
        if (collapseWhitespace(text) === collapseWhitespace(browserText)) return;
        e.clipboardData.setData("text/plain", text);
        e.preventDefault();
      },
      [ref, scale, pageHeight, page.structured]
    );

    const onMouseUp = useCallback(() => {
      if (!onSelectionPdfRects) return;
      const inner = typeof ref === "object" && ref ? ref.current : null;
      if (!inner) return;
      const sel = typeof window !== "undefined" ? window.getSelection() : null;
      if (!sel || sel.isCollapsed || sel.rangeCount === 0) return;
      const range = sel.getRangeAt(0);
      if (!inner.contains(range.commonAncestorContainer)) return;
      const innerRect = inner.getBoundingClientRect();
      if (innerRect.width === 0 || innerRect.height === 0) return;
      const clientRects = Array.from(range.getClientRects()).filter((r) => r.width > 0 && r.height > 0);
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
    }, [onSelectionPdfRects, scale, pageHeight, ref]);

    return (
      <div
        ref={ref}
        onMouseUp={onMouseUp}
        onCopy={onCopy}
        className="officeai-pdf-text-layer pointer-events-auto select-text"
        style={{
          width: rotatedW,
          height: rotatedH,
          // PDF.js's `TextLayer` reads `--scale-factor` for its
          // internal CSS calc()s. Without this every zoom would
          // require a full re-layout via `update()`; with it the
          // spans hot-reflow without rebuilding.
          ["--scale-factor" as unknown as string]: String(scale),
          ["--total-scale-factor" as unknown as string]: String(scale),
        }}
      />
    );
  }
);

function collapseWhitespace(s: string): string {
  return s.replace(/\s+/g, " ").trim();
}

/**
 * CSS transform that maps an un-rotated overlay box onto the
 * rotated visual frame. Mirrors PDF.js's viewport rotation so the
 * annotation / sticky-note overlays line up with the rasterised
 * page even when the user rotated the viewport.
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

/**
 * Internal handle returned by `mountOfficialTextLayer` so the
 * `useEffect` cleanup path can cancel an in-flight render.
 */
interface TextLayerHandle {
  cancel(): void;
}

/**
 * Construct PDF.js's canonical `TextLayer` against a container
 * element and the engine's native viewport. Returns a handle the
 * caller can cancel on unmount / rerender.
 */
async function mountOfficialTextLayer(opts: {
  enginePage: PdfEnginePage;
  container: HTMLElement;
  scale: number;
  rotation: PdfRotation;
}): Promise<TextLayerHandle> {
  const { enginePage, container, scale, rotation } = opts;
  // Lazy-load via dynamic import to keep the SSR bundle clean — the
  // `pdfjs-dist` ESM module imports browser-only globals at the top
  // level and can't be statically required by Next.js's server side.
  const pdfjs = (await import("pdfjs-dist/legacy/build/pdf.mjs")) as unknown as {
    TextLayer: new (params: { textContentSource: unknown; container: HTMLElement; viewport: unknown }) => {
      render(): Promise<unknown>;
      cancel(): void;
    };
  };
  if (typeof pdfjs.TextLayer !== "function") {
    throw new Error("pdfjs-dist: TextLayer export missing. Required pdfjs-dist >= 4.0.");
  }
  // Clear any previous content so re-mounting is safe.
  while (container.firstChild) container.removeChild(container.firstChild);
  const viewport = enginePage.getViewport({ scale, rotation });
  const textContentSource = await enginePage.getTextContentSource({
    includeMarkedContent: true,
  });
  const layer = new pdfjs.TextLayer({
    textContentSource,
    container,
    viewport: viewport.raw,
  });
  let cancelled = false;
  const renderPromise = layer.render();
  void renderPromise.catch(() => {
    /* swallow; cancellation surfaces here */
  });
  return {
    cancel: () => {
      if (cancelled) return;
      cancelled = true;
      try {
        layer.cancel();
      } catch {
        /* noop */
      }
    },
  };
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

function pageAnnotations(snapshot: PdfSnapshot, pageNumber: number): ReadonlyArray<PdfAnnotation> {
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
    <div className="pointer-events-none absolute inset-0" style={{ width: rotatedW, height: rotatedH }}>
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
                  <div className="whitespace-pre-wrap break-words">{a.contents || "(no contents)"}</div>
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
    <div className="absolute z-30" style={{ left: x, top: y }} onClick={(e) => e.stopPropagation()}>
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

function PdfMatchHighlight({
  highlight,
  pageWidth,
  pageHeight,
  scale,
  rotation,
}: {
  readonly highlight: PdfHighlight;
  readonly pageWidth: number;
  readonly pageHeight: number;
  readonly scale: number;
  readonly rotation: PdfRotation;
}): ReactNode {
  const quads = highlight.quads ?? [];
  if (quads.length > 0) {
    const innerWidth = pageWidth * scale;
    const innerHeight = pageHeight * scale;
    const rotatedW = rotation === 90 || rotation === 270 ? innerHeight : innerWidth;
    const rotatedH = rotation === 90 || rotation === 270 ? innerWidth : innerHeight;
    const transform = textLayerTransform(rotation, innerWidth, innerHeight);
    return (
      <div
        key={highlight.nonce}
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
          {quads.map((q, i) => {
            const [x1, y1, x2, y2] = q;
            const left = x1 * scale;
            const top = (pageHeight - y2) * scale;
            const width = Math.max(2, (x2 - x1) * scale);
            const height = Math.max(2, (y2 - y1) * scale);
            return (
              <div
                key={i}
                className="rounded-sm bg-[var(--accent-light)] ring-2 ring-[var(--accent)]/70"
                style={{ position: "absolute", left, top, width, height, mixBlendMode: "multiply" }}
              />
            );
          })}
        </div>
      </div>
    );
  }
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

/**
 * How a two-up row places its page(s) inside the spread:
 * - `right`   – single page on the right slot (front cover).
 * - `left`    – single page on the left slot (lone trailing page
 *               of an odd-page document).
 * - `spread`  – two pages, left + right, touching at the binding.
 */
type TwoUpSlot = "left" | "right" | "spread";

interface RenderRow {
  readonly key: string;
  readonly kind: "single" | "two-up";
  readonly pages: ReadonlyArray<PdfPage>;
  /** Only meaningful when `kind === "two-up"`. */
  readonly slot: TwoUpSlot;
}

function collectRenderTargets(
  pages: ReadonlyArray<PdfPage>,
  viewMode: PdfViewMode,
  currentPage: number
): ReadonlyArray<RenderRow> {
  if (viewMode === "single") {
    const p = pages.find((x) => x.pageNumber === currentPage) ?? pages[0];
    return p ? [{ key: `single-${p.id}`, kind: "single", pages: [p], slot: "spread" }] : [];
  }
  if (viewMode === "two-up") {
    const rows: RenderRow[] = [];
    // The cover (page 1) sits alone in the right slot — same
    // convention as a real book where the front cover is the
    // right-hand page when the book lies open.
    if (pages[0]) {
      rows.push({
        key: `cover-${pages[0].id}`,
        kind: "two-up",
        pages: [pages[0]],
        slot: "right",
      });
    }
    for (let i = 1; i < pages.length; i += 2) {
      const left = pages[i];
      if (!left) break;
      const right = pages[i + 1];
      if (right) {
        rows.push({
          key: `pair-${left.id}`,
          kind: "two-up",
          pages: [left, right],
          slot: "spread",
        });
      } else {
        // Odd page count → final page sits alone in the left slot
        // (back cover position) so the spread above stays aligned.
        rows.push({
          key: `pair-${left.id}`,
          kind: "two-up",
          pages: [left],
          slot: "left",
        });
      }
    }
    return rows;
  }
  return pages.map((p) => ({
    key: `cont-${p.id}`,
    kind: "single",
    pages: [p],
    slot: "spread",
  }));
}
