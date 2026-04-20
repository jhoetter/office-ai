"use client";

import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { ChevronDown, ChevronRight, FileText, ListTree, MessageSquare } from "lucide-react";
import type { PdfEngineDocument } from "@officeai/pdf-engine";
import type { PdfAnnotation, PdfComment, PdfOutlineNode, PdfPage, PdfSnapshot } from "@officeai/pdf";
import { useTranslator } from "@/lib/i18n";

export type PdfSidebarTab = "thumbnails" | "outline" | "annotations";

export interface PdfSidebarProps {
  readonly snapshot: PdfSnapshot | null;
  readonly engineDoc: PdfEngineDocument | null;
  readonly currentPage: number;
  readonly viewportRotation: 0 | 90 | 180 | 270;
  readonly tab: PdfSidebarTab;
  readonly onTabChange: (tab: PdfSidebarTab) => void;
  readonly onJumpToPage: (pageNumber: number) => void;
  /**
   * Optional onClick handler when a comment is picked. The editor
   * uses this to drive the canvas highlight + page jump in one go.
   */
  readonly onJumpToComment?: (commentId: string) => void;
  /**
   * Optional onClick handler for an annotation entry. Defaults to
   * `onJumpToPage(annotation.pageNumber)` when omitted.
   */
  readonly onJumpToAnnotation?: (annotation: PdfAnnotation) => void;
}

/**
 * Three-tab left sidebar for the PDF viewer.
 *
 *   - **Thumbnails** — virtualised list of low-resolution page
 *     renders. Each thumbnail lazy-renders into its own canvas via
 *     `IntersectionObserver`, so opening a 500-page brief doesn't
 *     melt the worker queue.
 *   - **Outline** — recursive tree built from `snapshot.root.outline`.
 *     Click a node to jump; nodes without a page destination still
 *     render but as a non-interactive entry (link annotations target
 *     external URIs in that case, surfaced with an `↗` glyph).
 *   - **Annotations** — flat list grouped by page, including the
 *     comment threads. Click-to-jump anchors back to the canvas.
 *
 * The sidebar reads the snapshot directly — no extra subscription —
 * because the editor already re-renders on every snapshot tick and
 * passes a fresh `snapshot` reference. We deliberately avoid
 * memoising on the snapshot identity (`useMemo(...,[snapshot])`)
 * because tabs that aren't open don't need the work; the per-tab
 * components own their own derivations and are mounted only when
 * active.
 */
export function PdfSidebar(props: PdfSidebarProps): ReactNode {
  const { snapshot, engineDoc, currentPage, viewportRotation, tab, onTabChange, onJumpToPage } = props;
  const { t } = useTranslator();
  return (
    <aside
      data-testid="pdf-sidebar"
      className="hidden w-[220px] shrink-0 flex-col overflow-hidden border-r border-divider bg-surface md:flex"
    >
      <div className="flex shrink-0 border-b border-divider">
        <SidebarTabButton
          active={tab === "thumbnails"}
          icon={<FileText size={14} />}
          label={t("pdf.thumbnails")}
          testId="pdf-sidebar-tab-thumbnails"
          onClick={() => onTabChange("thumbnails")}
        />
        <SidebarTabButton
          active={tab === "outline"}
          icon={<ListTree size={14} />}
          label={t("pdf.outline")}
          testId="pdf-sidebar-tab-outline"
          onClick={() => onTabChange("outline")}
        />
        <SidebarTabButton
          active={tab === "annotations"}
          icon={<MessageSquare size={14} />}
          label={t("pdf.annotations")}
          testId="pdf-sidebar-tab-annotations"
          onClick={() => onTabChange("annotations")}
        />
      </div>
      <div className="min-h-0 flex-1 overflow-y-auto">
        {tab === "thumbnails" ? (
          <ThumbnailsPanel
            snapshot={snapshot}
            engineDoc={engineDoc}
            currentPage={currentPage}
            viewportRotation={viewportRotation}
            onJumpToPage={onJumpToPage}
          />
        ) : tab === "outline" ? (
          <OutlinePanel snapshot={snapshot} currentPage={currentPage} onJumpToPage={onJumpToPage} />
        ) : (
          <AnnotationsPanel
            snapshot={snapshot}
            onJumpToPage={onJumpToPage}
            {...(props.onJumpToComment ? { onJumpToComment: props.onJumpToComment } : {})}
            {...(props.onJumpToAnnotation ? { onJumpToAnnotation: props.onJumpToAnnotation } : {})}
          />
        )}
      </div>
    </aside>
  );
}

interface SidebarTabButtonProps {
  readonly active: boolean;
  readonly icon: ReactNode;
  readonly label: string;
  readonly testId: string;
  readonly onClick: () => void;
}

function SidebarTabButton({ active, icon, label, testId, onClick }: SidebarTabButtonProps): ReactNode {
  return (
    <button
      type="button"
      onClick={onClick}
      data-testid={testId}
      aria-pressed={active}
      title={label}
      className={
        "flex flex-1 items-center justify-center gap-1.5 px-2 py-2 text-[11px] uppercase tracking-wide " +
        (active
          ? "border-b-2 border-[var(--accent)] text-primary"
          : "border-b-2 border-transparent text-tertiary hover:text-primary")
      }
    >
      {icon}
      <span className="hidden xl:inline">{label}</span>
    </button>
  );
}

// ─────────────────────────────────────────────────────────────────────
// Thumbnails

interface ThumbnailsPanelProps {
  readonly snapshot: PdfSnapshot | null;
  readonly engineDoc: PdfEngineDocument | null;
  readonly currentPage: number;
  readonly viewportRotation: 0 | 90 | 180 | 270;
  readonly onJumpToPage: (pageNumber: number) => void;
}

const THUMB_SCALE = 0.18;
const THUMB_MAX_WIDTH = 170;

function ThumbnailsPanel({
  snapshot,
  engineDoc,
  currentPage,
  viewportRotation,
  onJumpToPage,
}: ThumbnailsPanelProps): ReactNode {
  const { t } = useTranslator();
  const pages = snapshot?.root.pages ?? [];

  // Keep the active thumbnail roughly centred in the rail whenever
  // `currentPage` moves — including when the user is scrolling the
  // main canvas. Centring (rather than `block: "nearest"`) means the
  // reader can always see the previous + next page above and below
  // the highlighted one, so the rail acts as a "what's coming"
  // preview while reading. The effect only reacts to `currentPage`
  // changes, never to sidebar scroll, so the main canvas stays put
  // when the user scrubs the rail by hand.
  const containerRef = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    const inner = containerRef.current;
    if (!inner) return;
    const target = inner.querySelector<HTMLElement>(`[data-thumb-page="${currentPage}"]`);
    if (!target) return;
    // The scrollable element is the panel wrapper one level up from
    // our flex column (see `PdfSidebar`'s `overflow-y-auto` div).
    // Walk up until we find an actually-scrollable ancestor so the
    // visibility check uses the right viewport.
    let scroller: HTMLElement | null = inner.parentElement;
    while (scroller && scroller.scrollHeight <= scroller.clientHeight) {
      scroller = scroller.parentElement;
    }
    if (!scroller) return;
    const tTop = target.offsetTop - scroller.offsetTop;
    const tBot = tTop + target.offsetHeight;
    const visTop = scroller.scrollTop;
    const visBot = visTop + scroller.clientHeight;
    const fullyVisible = tTop >= visTop && tBot <= visBot;
    // Always re-centre when partially / fully off-screen; for an
    // already fully visible thumbnail we still nudge it back toward
    // the middle so the next page peeks into view as the reader
    // scrolls forward.
    const desired = target.offsetTop - scroller.offsetTop - (scroller.clientHeight - target.offsetHeight) / 2;
    const clamped = Math.max(0, Math.min(scroller.scrollHeight - scroller.clientHeight, desired));
    if (!fullyVisible || Math.abs(clamped - visTop) > target.offsetHeight) {
      scroller.scrollTo({ top: clamped, behavior: "smooth" });
    }
  }, [currentPage]);

  if (!snapshot || pages.length === 0) {
    return <p className="p-4 text-xs text-tertiary">{t("pdf.loading")}</p>;
  }

  return (
    <div ref={containerRef} className="flex flex-col items-center gap-2 px-2 py-3">
      {pages.map((p) => (
        <ThumbnailItem
          key={p.id}
          page={p}
          engineDoc={engineDoc}
          rotation={viewportRotation}
          active={p.pageNumber === currentPage}
          onClick={() => onJumpToPage(p.pageNumber)}
        />
      ))}
    </div>
  );
}

interface ThumbnailItemProps {
  readonly page: PdfPage;
  readonly engineDoc: PdfEngineDocument | null;
  readonly rotation: 0 | 90 | 180 | 270;
  readonly active: boolean;
  readonly onClick: () => void;
}

function ThumbnailItem({ page, engineDoc, rotation, active, onClick }: ThumbnailItemProps): ReactNode {
  const ref = useRef<HTMLDivElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [visible, setVisible] = useState(false);
  const [renderedAt, setRenderedAt] = useState<number | null>(null);

  // Derive a CSS size that respects the page aspect ratio (with the
  // viewport rotation applied) while capping the width at
  // THUMB_MAX_WIDTH. The strip mirrors the main canvas so a 90°
  // landscape spin in the toolbar lands in the rail too.
  const intrinsicW = rotation === 90 || rotation === 270 ? page.height : page.width;
  const intrinsicH = rotation === 90 || rotation === 270 ? page.width : page.height;
  const widthCss = Math.min(THUMB_MAX_WIDTH, Math.round(intrinsicW * THUMB_SCALE * 4));
  const heightCss = Math.round((widthCss * intrinsicH) / intrinsicW);

  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const io = new IntersectionObserver(
      (entries) => {
        for (const e of entries) {
          if (e.isIntersecting) {
            setVisible(true);
            io.disconnect();
            return;
          }
        }
      },
      { rootMargin: "200px 0px" }
    );
    io.observe(el);
    return () => io.disconnect();
  }, []);

  useEffect(() => {
    if (!visible || !engineDoc) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;
    let cancelled = false;
    let enginePage: import("@officeai/pdf-engine").PdfEnginePage | null = null;
    void (async () => {
      try {
        enginePage = await engineDoc.getPage(page.pageNumber);
        if (cancelled) return;
        const dpr = typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1;
        canvas.width = Math.max(1, Math.floor(widthCss * dpr));
        canvas.height = Math.max(1, Math.floor(heightCss * dpr));
        canvas.style.width = `${widthCss}px`;
        canvas.style.height = `${heightCss}px`;
        ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
        ctx.fillStyle = "white";
        ctx.fillRect(0, 0, widthCss, heightCss);
        const renderScale = (widthCss / intrinsicW) * dpr;
        await enginePage.render({ canvas, scale: renderScale, rotation });
        if (cancelled) return;
        setRenderedAt(Date.now());
      } catch {
        // Thumbnail failures are silent — the placeholder stays
        // visible. We don't want a broken page to fill the strip
        // with red error text and distract from the main reading
        // surface.
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
  }, [visible, engineDoc, page.pageNumber, intrinsicW, widthCss, heightCss, rotation]);

  return (
    <div ref={ref} data-thumb-page={page.pageNumber} className="flex flex-col items-center gap-1">
      <button
        type="button"
        onClick={onClick}
        title={`Page ${page.pageNumber}`}
        data-testid={`pdf-thumbnail-${page.pageNumber}`}
        className={
          "relative flex items-center justify-center overflow-hidden rounded border bg-white text-xs text-tertiary " +
          (active
            ? "border-[var(--accent)] ring-2 ring-[var(--accent)]/40"
            : "border-divider hover:border-[var(--accent)]/60")
        }
        style={{ width: widthCss, height: heightCss }}
      >
        <canvas ref={canvasRef} className="block h-full w-full" />
        {!renderedAt ? <span className="absolute text-[10px]">{page.pageNumber}</span> : null}
      </button>
      <span className="text-[10px] tabular-nums text-tertiary">{page.pageNumber}</span>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────
// Outline

interface OutlinePanelProps {
  readonly snapshot: PdfSnapshot | null;
  readonly currentPage: number;
  readonly onJumpToPage: (pageNumber: number) => void;
}

function OutlinePanel({ snapshot, currentPage, onJumpToPage }: OutlinePanelProps): ReactNode {
  const { t } = useTranslator();
  const outline = snapshot?.root.outline ?? [];

  // Resolve which outline node represents the user's current page.
  // We walk the tree in document order and pick the *last* entry
  // whose destination page is ≤ currentPage — i.e. the deepest
  // section the reader has scrolled into. This mirrors how Preview
  // / Acrobat highlight the active heading.
  const { activeId, ancestorIds } = useMemo(
    () => resolveActiveOutline(outline, currentPage),
    [outline, currentPage]
  );

  // Per-node controlled open state. Defaults expand the top level
  // and any ancestor of the active node, so navigating into a deep
  // chapter automatically opens the reader's path through the tree.
  const [openMap, setOpenMap] = useState<Record<string, boolean>>({});
  useEffect(() => {
    if (ancestorIds.length === 0) return;
    setOpenMap((prev) => {
      let next = prev;
      for (const id of ancestorIds) {
        if (!next[id]) {
          if (next === prev) next = { ...prev };
          next[id] = true;
        }
      }
      return next;
    });
  }, [ancestorIds]);
  const onToggle = useCallback((id: string) => {
    setOpenMap((prev) => ({ ...prev, [id]: !prev[id] }));
  }, []);

  if (outline.length === 0) {
    return <p className="p-4 text-xs text-tertiary">{t("pdf.noOutline")}</p>;
  }
  return (
    <ul className="px-1 py-2 text-xs" role="tree">
      {outline.map((node) => (
        <OutlineEntry
          key={node.id}
          node={node}
          depth={0}
          activeId={activeId}
          ancestorIds={ancestorIds}
          openMap={openMap}
          onToggle={onToggle}
          onJumpToPage={onJumpToPage}
        />
      ))}
    </ul>
  );
}

interface OutlineEntryProps {
  readonly node: PdfOutlineNode;
  readonly depth: number;
  readonly activeId: string | null;
  readonly ancestorIds: ReadonlyArray<string>;
  readonly openMap: Readonly<Record<string, boolean>>;
  readonly onToggle: (id: string) => void;
  readonly onJumpToPage: (pageNumber: number) => void;
}

function OutlineEntry({
  node,
  depth,
  activeId,
  ancestorIds,
  openMap,
  onToggle,
  onJumpToPage,
}: OutlineEntryProps): ReactNode {
  const hasChildren = node.children.length > 0;
  const isActive = activeId === node.id;
  const isAncestor = ancestorIds.includes(node.id);
  // Default: expand top-level + any ancestor of the active entry.
  // The controlled openMap overrides — once a user clicks a chevron
  // we honour their choice and stop auto-expanding.
  const explicit = openMap[node.id];
  const open = explicit !== undefined ? explicit : depth < 1 || isAncestor;
  const labelRef = useRef<HTMLButtonElement | null>(null);
  // Scroll the active entry into view when the user navigates
  // outside the sidebar (page change via canvas scroll, toolbar,
  // etc.). `block: "nearest"` avoids jumpy auto-scroll when the
  // entry is already visible.
  useEffect(() => {
    if (!isActive) return;
    const el = labelRef.current;
    if (!el) return;
    el.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }, [isActive]);

  const handleClick = useCallback((): void => {
    if (typeof node.pageNumber === "number") {
      onJumpToPage(node.pageNumber);
    } else if (node.uri && typeof window !== "undefined") {
      window.open(node.uri, "_blank", "noopener,noreferrer");
    }
  }, [node.pageNumber, node.uri, onJumpToPage]);
  const labelClass =
    "flex flex-1 items-baseline justify-between gap-2 truncate text-left rounded px-1 py-0.5 " +
    (isActive
      ? "bg-[var(--accent-light)] font-medium text-[var(--accent)]"
      : "text-primary hover:text-[var(--accent)]");
  return (
    <li
      role="treeitem"
      aria-expanded={hasChildren ? open : undefined}
      aria-current={isActive ? "page" : undefined}
    >
      <div
        className="group flex items-center gap-1 rounded hover:bg-hover"
        style={{ paddingLeft: 4 + depth * 12 }}
      >
        {hasChildren ? (
          <button
            type="button"
            onClick={() => onToggle(node.id)}
            className="flex h-4 w-4 items-center justify-center text-tertiary"
            title={open ? "Collapse" : "Expand"}
          >
            {open ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
          </button>
        ) : (
          <span className="inline-block h-4 w-4" aria-hidden />
        )}
        <button
          ref={labelRef}
          type="button"
          onClick={handleClick}
          className={labelClass}
          title={node.title}
          data-testid={`pdf-outline-${node.id}`}
          data-active={isActive ? "true" : undefined}
        >
          <span className="truncate">{node.title}</span>
          {typeof node.pageNumber === "number" ? (
            <span
              className={
                "shrink-0 text-[10px] tabular-nums " + (isActive ? "text-[var(--accent)]" : "text-tertiary")
              }
            >
              {node.pageNumber}
            </span>
          ) : node.uri ? (
            <span className="shrink-0 text-[10px] text-tertiary" aria-hidden>
              ↗
            </span>
          ) : null}
        </button>
      </div>
      {hasChildren && open ? (
        <ul role="group">
          {node.children.map((child) => (
            <OutlineEntry
              key={child.id}
              node={child}
              depth={depth + 1}
              activeId={activeId}
              ancestorIds={ancestorIds}
              openMap={openMap}
              onToggle={onToggle}
              onJumpToPage={onJumpToPage}
            />
          ))}
        </ul>
      ) : null}
    </li>
  );
}

/**
 * Walk the outline in document order and pick the deepest entry
 * whose destination page is ≤ `currentPage`. Returns both the
 * winning node id and the chain of ancestor ids leading to it so
 * the panel can auto-expand the reader's path.
 *
 * Why "deepest": if the reader is on page 42 and the outline says
 * Chapter 3 starts on page 40 with section 3.2 on page 42, we want
 * the active entry to be 3.2 — that's where the user is reading,
 * not the chapter header.
 */
function resolveActiveOutline(
  outline: ReadonlyArray<PdfOutlineNode>,
  currentPage: number
): { activeId: string | null; ancestorIds: ReadonlyArray<string> } {
  let bestId: string | null = null;
  let bestPage = -Infinity;
  let bestPath: ReadonlyArray<string> = [];
  const visit = (nodes: ReadonlyArray<PdfOutlineNode>, path: ReadonlyArray<string>): void => {
    for (const node of nodes) {
      if (
        typeof node.pageNumber === "number" &&
        node.pageNumber <= currentPage &&
        node.pageNumber >= bestPage
      ) {
        bestId = node.id;
        bestPage = node.pageNumber;
        bestPath = path;
      }
      if (node.children.length > 0) visit(node.children, [...path, node.id]);
    }
  };
  visit(outline, []);
  return { activeId: bestId, ancestorIds: bestPath };
}

// ─────────────────────────────────────────────────────────────────────
// Annotations

interface AnnotationsPanelProps {
  readonly snapshot: PdfSnapshot | null;
  readonly onJumpToPage: (pageNumber: number) => void;
  readonly onJumpToComment?: (commentId: string) => void;
  readonly onJumpToAnnotation?: (annotation: PdfAnnotation) => void;
}

interface PageGroup {
  readonly pageNumber: number;
  readonly annotations: ReadonlyArray<PdfAnnotation>;
  readonly comments: ReadonlyArray<PdfComment>;
}

function AnnotationsPanel({
  snapshot,
  onJumpToPage,
  onJumpToComment,
  onJumpToAnnotation,
}: AnnotationsPanelProps): ReactNode {
  const { t } = useTranslator();
  const groups = useMemo<ReadonlyArray<PageGroup>>(() => {
    if (!snapshot) return [];
    const annoByPage = new Map<number, PdfAnnotation[]>();
    for (const a of snapshot.root.annotations) {
      const list = annoByPage.get(a.pageNumber) ?? [];
      list.push(a);
      annoByPage.set(a.pageNumber, list);
    }
    const commentsByPage = new Map<number, PdfComment[]>();
    for (const c of snapshot.root.comments) {
      // Reply comments are surfaced underneath their root in the
      // shared CommentsSidebar; here we only list root threads to
      // keep the tab readable.
      if (c.parentId) continue;
      const list = commentsByPage.get(c.pageNumber) ?? [];
      list.push(c);
      commentsByPage.set(c.pageNumber, list);
    }
    const allPages = new Set<number>([...annoByPage.keys(), ...commentsByPage.keys()]);
    const sorted = [...allPages].sort((a, b) => a - b);
    return sorted.map((pageNumber) => ({
      pageNumber,
      annotations: annoByPage.get(pageNumber) ?? [],
      comments: commentsByPage.get(pageNumber) ?? [],
    }));
  }, [snapshot]);

  if (groups.length === 0) {
    return <p className="p-4 text-xs text-tertiary">{t("pdf.noAnnotations")}</p>;
  }

  return (
    <div className="flex flex-col gap-3 p-3 text-xs">
      {groups.map((g) => (
        <section key={g.pageNumber} className="rounded border border-divider bg-background">
          <header className="flex items-center justify-between border-b border-divider px-2 py-1">
            <button
              type="button"
              onClick={() => onJumpToPage(g.pageNumber)}
              className="text-[11px] font-semibold text-primary hover:text-[var(--accent)]"
              data-testid={`pdf-annotations-page-${g.pageNumber}`}
            >
              {t("pdf.page")} {g.pageNumber}
            </button>
            <span className="text-[10px] tabular-nums text-tertiary">
              {g.annotations.length + g.comments.length}
            </span>
          </header>
          <ul className="divide-y divide-divider">
            {g.annotations.map((a) => (
              <li key={a.id} className="px-2 py-1.5">
                <button
                  type="button"
                  className="flex w-full items-baseline justify-between gap-2 text-left text-primary hover:text-[var(--accent)]"
                  onClick={() => (onJumpToAnnotation ? onJumpToAnnotation(a) : onJumpToPage(a.pageNumber))}
                  title={a.contents ?? a.kind}
                  data-testid={`pdf-annotation-${a.id}`}
                >
                  <span className="truncate">{summariseAnnotation(a)}</span>
                  <span className="shrink-0 text-[10px] uppercase text-tertiary">{a.kind}</span>
                </button>
              </li>
            ))}
            {g.comments.map((c) => (
              <li key={c.id} className="px-2 py-1.5">
                <button
                  type="button"
                  className="flex w-full flex-col gap-0.5 text-left hover:text-[var(--accent)]"
                  onClick={() => (onJumpToComment ? onJumpToComment(c.id) : onJumpToPage(c.pageNumber))}
                  data-testid={`pdf-comment-${c.id}`}
                >
                  <span className="text-[11px] font-medium text-primary">
                    {c.author}
                    {c.resolved ? (
                      <span className="ml-1 text-[10px] uppercase text-tertiary">resolved</span>
                    ) : null}
                  </span>
                  <span className="line-clamp-2 text-[11px] text-secondary">{c.text}</span>
                </button>
              </li>
            ))}
          </ul>
        </section>
      ))}
    </div>
  );
}

function summariseAnnotation(a: PdfAnnotation): string {
  if (a.contents && a.contents.trim().length > 0) {
    return a.contents.trim().slice(0, 80);
  }
  if (a.kind === "link") {
    if (a.url) return a.url;
    if (a.destPage) return `→ Page ${a.destPage}`;
  }
  return a.subtype || a.kind;
}
