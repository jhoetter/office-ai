"use client";

import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import type { PdfEngineDocument, PdfEnginePage, PdfEngineTextItem } from "@officeai/pdf-engine";
import type { PdfPage, PdfRotation, PdfSnapshot } from "@officeai/pdf";
import { useTranslator } from "@/lib/i18n";
import { applyDarkModeToContext, darkModeCssFilter } from "./darkMode";

/** Canvas-side view modes — 1:1 with the toolbar enum. */
export type PdfViewMode = "single" | "continuous" | "two-up";

/**
 * Smart-invert dark mode mode picker. The toolbar lets the user
 * pick between the cheap CSS filter (instant, photo-inverting),
 * the per-pixel pass that preserves photographs (~5 ms / page on
 * an A4 at scale 1.5), or `off`.
 */
export type PdfDarkModeStrategy = "off" | "css" | "smart";

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
  } = props;
  const { t } = useTranslator();
  const containerRef = useRef<HTMLDivElement | null>(null);
  const [containerWidth, setContainerWidth] = useState(0);

  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => {
      setContainerWidth(el.clientWidth);
    });
    ro.observe(el);
    setContainerWidth(el.clientWidth);
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
  // for a different page. Skip when we're already there.
  useEffect(() => {
    const root = containerRef.current;
    if (!root) return;
    if (viewMode === "single") return;
    const target = root.querySelector<HTMLElement>(
      `[data-page-number="${currentPage}"]`
    );
    if (!target) return;
    const top = target.offsetTop;
    if (Math.abs(top - root.scrollTop) < 4) return;
    root.scrollTo({ top, behavior: "smooth" });
  }, [currentPage, viewMode, scale]);

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
  } = props;
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
        await enginePage.render({ canvas, scale: scale * dpr });
        if (cancelled || token !== renderTokenRef.current) return;

        if (darkMode === "smart") {
          ctx.setTransform(1, 0, 0, 1, 0, 0);
          applyDarkModeToContext(ctx, canvas.width, canvas.height);
        }

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
  }, [visible, engineDoc, page.pageNumber, scale, cssWidth, cssHeight, darkMode]);

  return (
    <div
      data-page-number={page.pageNumber}
      data-testid={`pdf-page-${page.pageNumber}`}
      onClick={onClick}
      className={
        "relative cursor-default rounded-md border bg-white shadow-sm " +
        (active ? "border-[var(--accent)]" : "border-divider")
      }
      style={{ width: cssWidth, height: cssHeight }}
    >
      <canvas
        ref={canvasRef}
        className="block h-full w-full rounded-md"
        style={{
          filter: darkMode === "css" ? darkModeCssFilter(true) : "none",
          transform: viewportRotation === 0 ? undefined : `rotate(${viewportRotation}deg)`,
        }}
        aria-label={t("pdf.pageOf", { n: page.pageNumber, total: totalPages })}
      />
      {visible && textItems.length > 0 ? (
        <PdfTextLayer items={textItems} scale={scale} pageHeight={page.height} />
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
  readonly pageHeight: number;
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
 */
function PdfTextLayer({ items, scale, pageHeight }: PdfTextLayerProps): ReactNode {
  return (
    <div
      aria-hidden
      className="pointer-events-auto absolute inset-0 select-text overflow-hidden text-transparent"
      style={{ lineHeight: 1 }}
    >
      {items.map((it, idx) => {
        if (!it.str) return null;
        const tx = it.transform;
        const fontHeight = Math.abs(tx[3]) || it.height || 1;
        const left = tx[4] * scale;
        const top = (pageHeight - tx[5] - fontHeight) * scale;
        const width = it.width * scale;
        const height = fontHeight * scale;
        return (
          <span
            key={idx}
            style={{
              position: "absolute",
              left,
              top,
              width,
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
