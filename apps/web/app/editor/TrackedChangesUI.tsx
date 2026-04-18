"use client";

import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from "react";
import { Check, X } from "lucide-react";
import { cn } from "@officeai/ui";
import { collectRevisionsWithPreview, snippet } from "@/lib/format-helpers";
import type { DocxSnapshot } from "@officeai/docx";

/**
 * Word-style "All Markup" surface for tracked changes.
 *
 * The component is split into two pieces so each can be mounted at the
 * spot where its DOM dependencies actually live:
 *
 *  - {@link TrackedChangesHover} attaches a delegated `mouseover`
 *    handler to the editor host so a small Accept/Reject popover
 *    follows the cursor over `.pm-revision-ins` underlines. Mount
 *    once near the top of `DocxEditor`; safe to keep mounted even
 *    when the document has no revisions (the listener is a no-op).
 *
 *  - {@link TrackedChangesMargin} paints the WORD-style review chrome
 *    inside the editor's scroll container: a vertical change bar in
 *    the LEFT page margin per revision row, and a stack of
 *    Accept/Reject balloons in the RIGHT gutter labeled with the
 *    revision author. Mount as a sibling of the page card inside the
 *    scroll container so absolute positioning is anchored correctly
 *    and balloons pan with the document on scroll.
 *
 * The balloon copy mirrors Word in German ("hat gelöscht …" /
 * "hat eingefügt …") so the UI matches the screenshot the design
 * brief is anchored to. The substantive content (author, preview
 * text) flows out of the snapshot via `collectRevisionsWithPreview`.
 */

export function TrackedChangesHover(props: {
  editorHost: HTMLElement | null;
  onAccept: (revisionId: string) => Promise<void> | void;
  onReject: (revisionId: string) => Promise<void> | void;
}): ReactNode {
  return <InlineHoverWidget {...props} />;
}

function InlineHoverWidget(props: {
  editorHost: HTMLElement | null;
  onAccept: (id: string) => Promise<void> | void;
  onReject: (id: string) => Promise<void> | void;
}): ReactNode {
  const [hovered, setHovered] = useState<{
    revisionId: string;
    rect: DOMRect;
    type: "ins" | "del";
  } | null>(null);
  const widgetRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    const host = props.editorHost;
    if (!host) return;

    let hideTimer: ReturnType<typeof setTimeout> | null = null;
    const cancelHide = () => {
      if (hideTimer) {
        clearTimeout(hideTimer);
        hideTimer = null;
      }
    };
    const scheduleHide = () => {
      cancelHide();
      hideTimer = setTimeout(() => setHovered(null), 200);
    };

    const onOver = (e: MouseEvent) => {
      const target = e.target as HTMLElement | null;
      if (!target) return;
      // Deletions are zero-width inline placeholders (see globals.css)
      // so they are not addressable via hover; the right-margin
      // balloon is the deletion's interaction surface.
      const span = target.closest<HTMLElement>(".pm-revision-ins");
      if (!span) return;
      const revisionId = span.getAttribute("data-revision-id") ?? "";
      if (!revisionId) return;
      cancelHide();
      setHovered({
        revisionId,
        rect: span.getBoundingClientRect(),
        type: "ins",
      });
    };

    const onOut = (e: MouseEvent) => {
      const related = e.relatedTarget as Node | null;
      if (related && widgetRef.current?.contains(related)) return;
      scheduleHide();
    };

    host.addEventListener("mouseover", onOver);
    host.addEventListener("mouseout", onOut);
    return () => {
      cancelHide();
      host.removeEventListener("mouseover", onOver);
      host.removeEventListener("mouseout", onOut);
    };
  }, [props.editorHost]);

  if (!hovered) return null;

  const top = hovered.rect.top - 32 + window.scrollY;
  const left = hovered.rect.left + window.scrollX;

  return (
    <div
      ref={widgetRef}
      role="group"
      aria-label="Tracked change actions"
      className="tracked-change-widget pointer-events-auto fixed z-50 flex items-center gap-1 rounded-md border border-divider bg-surface px-1 py-0.5 text-xs shadow-md"
      style={{ top, left, position: "absolute" }}
    >
      <span aria-hidden className="px-1 text-[10px] font-medium uppercase text-[var(--accent)]">
        Insertion
      </span>
      <button
        type="button"
        title="Accept change"
        aria-label="Accept change"
        onClick={() => void props.onAccept(hovered.revisionId)}
        className="rounded p-1 text-[var(--success)] hover:bg-[var(--success)]/10"
      >
        <Check size={12} />
      </button>
      <button
        type="button"
        title="Reject change"
        aria-label="Reject change"
        onClick={() => void props.onReject(hovered.revisionId)}
        className="rounded p-1 text-[var(--error)] hover:bg-[var(--error)]/10"
      >
        <X size={12} />
      </button>
    </div>
  );
}

interface BalloonLayout {
  readonly revisionId: string;
  readonly revisionType: "ins" | "del";
  readonly author: string;
  readonly previewText: string;
  /** Top of the change in the scroll container's content coordinates. */
  readonly anchorTop: number;
  /** Final balloon top after collision-resolution stacking. */
  readonly balloonTop: number;
  /** Approximate vertical extent of the change for the left change-bar. */
  readonly anchorHeight: number;
}

export interface TrackedChangesMarginProps {
  snapshot: DocxSnapshot | null;
  /**
   * The PM editor host (`prose-pm` div). The component reads
   * `data-revision-id` spans inside this element to anchor balloons,
   * and locates the inner `.ProseMirror` page card to align the
   * change bar / balloon column.
   */
  editorHost: HTMLElement | null;
  /**
   * The scroll container that wraps both the page card and the
   * balloon overlay. Absolute positions are computed in this
   * element's content coordinate space so balloons stay aligned with
   * the page on scroll without a scroll listener.
   */
  scrollContainer: HTMLElement | null;
  onAccept: (revisionId: string) => Promise<void> | void;
  onReject: (revisionId: string) => Promise<void> | void;
}

const BALLOON_WIDTH = 220;
const BALLOON_GAP_PX = 6;
const BALLOON_GUTTER_PX = 16;
const CHANGEBAR_INSET_PX = 8;
const ESTIMATED_BALLOON_HEIGHT_PX = 64;

/**
 * Word-style margin chrome. Renders left-margin change bars and a
 * right-margin balloon column overlaid on the scroll container.
 *
 * Layout strategy:
 *  - For each revision span we read `getBoundingClientRect` and
 *    translate it into the scroll container's content coordinate
 *    space (`rect.top - container.top + container.scrollTop`). The
 *    same translation is applied to the page card so we know where
 *    to place the change-bar (just inside the page's left margin)
 *    and where the balloon column should start (just outside the
 *    page's right edge).
 *  - Balloons are stacked top-to-bottom using a one-pass
 *    collision-resolution: if a balloon would overlap its
 *    predecessor, it is pushed down so the two never overlap. The
 *    rendered height is read after the first paint via
 *    `ResizeObserver` and fed back into a re-layout, which keeps the
 *    stacking accurate even when a long preview wraps onto extra
 *    lines.
 *  - We re-measure on every snapshot change (revision content / DOM
 *    layout shifts) and on `ResizeObserver` ticks for the editor host
 *    (zoom, window resize, page-card width changes).
 */
export function TrackedChangesMargin(props: TrackedChangesMarginProps): ReactNode {
  const { snapshot, editorHost, scrollContainer, onAccept, onReject } = props;
  const [layouts, setLayouts] = useState<ReadonlyArray<BalloonLayout>>([]);
  const [columnLeft, setColumnLeft] = useState<number | null>(null);
  const [barLeft, setBarLeft] = useState<number | null>(null);
  const balloonRefs = useRef(new Map<string, HTMLElement>());
  // Bumped to trigger a re-measure pass after balloon DOM heights
  // settle from their initial paint.
  const [measureTick, setMeasureTick] = useState(0);

  const revisions = snapshot ? collectRevisionsWithPreview(snapshot) : [];
  const revisionsKey = revisions.map((r) => `${r.revisionId}:${r.revisionType}`).join("|");

  useLayoutEffect(() => {
    if (!editorHost || !scrollContainer) {
      setLayouts([]);
      setColumnLeft(null);
      setBarLeft(null);
      return;
    }
    // The .ProseMirror element is created lazily by mountDocxEditor;
    // resolve it on each pass rather than as a prop so we don't need
    // the parent to thread an extra ref.
    const pageCard = editorHost.querySelector<HTMLElement>(".ProseMirror");
    if (!pageCard) {
      setLayouts([]);
      setColumnLeft(null);
      setBarLeft(null);
      return;
    }
    if (revisions.length === 0) {
      setLayouts([]);
      return;
    }

    const containerRect = scrollContainer.getBoundingClientRect();
    const cardRect = pageCard.getBoundingClientRect();
    const scrollLeft = scrollContainer.scrollLeft;
    const scrollTop = scrollContainer.scrollTop;

    setColumnLeft(cardRect.right - containerRect.left + scrollLeft + BALLOON_GUTTER_PX);
    setBarLeft(cardRect.left - containerRect.left + scrollLeft + CHANGEBAR_INSET_PX);

    const next: BalloonLayout[] = [];
    let cursor = 0;
    for (const rev of revisions) {
      const span = editorHost.querySelector<HTMLElement>(
        `.pm-revision[data-revision-id="${cssEscape(rev.revisionId)}"]`
      );
      if (!span) continue;
      const rect = span.getBoundingClientRect();
      // Deletions are styled `font-size: 0` so they take zero room
      // inline (the page reads as if applied). Their rect collapses
      // to zero width AND zero height but still reports a meaningful
      // `top` for the line they sit on — except that some browsers
      // (Safari) report `top: 0` for fully collapsed inline spans.
      // Walk up to the nearest paragraph element so we always have a
      // line-sized fallback rect to anchor the change bar / balloon.
      const lineHost =
        rect.width === 0 && rect.height === 0
          ? (span.closest<HTMLElement>("p, h1, h2, h3, h4, h5, h6, li") ??
            span.parentElement ??
            span)
          : span;
      const anchorRect = lineHost === span ? rect : lineHost.getBoundingClientRect();
      // If even the parent has no rect (detached), there's nothing
      // sensible to anchor to.
      if (anchorRect.width === 0 && anchorRect.height === 0) continue;
      const anchorTop = anchorRect.top - containerRect.top + scrollTop;
      const anchorHeight = Math.max(anchorRect.height, 14);
      const measured =
        balloonRefs.current.get(rev.revisionId)?.getBoundingClientRect().height ??
        ESTIMATED_BALLOON_HEIGHT_PX;
      const desiredTop = anchorTop;
      const balloonTop = Math.max(desiredTop, cursor);
      cursor = balloonTop + measured + BALLOON_GAP_PX;
      next.push({
        revisionId: rev.revisionId,
        revisionType: rev.revisionType,
        author: rev.author,
        previewText: rev.previewText,
        anchorTop,
        balloonTop,
        anchorHeight,
      });
    }
    setLayouts(next);
    // Drop refs for revisions that no longer exist so the map doesn't
    // hold onto detached DOM nodes across edits.
    for (const id of Array.from(balloonRefs.current.keys())) {
      if (!next.some((l) => l.revisionId === id)) {
        balloonRefs.current.delete(id);
      }
    }
  }, [editorHost, scrollContainer, revisionsKey, snapshot, measureTick]);

  // Trigger a second pass once balloon heights are known so the
  // collision-stacking accounts for wrapped previews that exceed the
  // initial estimate. The dependency on `layouts.length` keeps this
  // bounded — we only re-tick when the number of balloons changes.
  useEffect(() => {
    if (layouts.length === 0) return;
    let raf = 0;
    raf = window.requestAnimationFrame(() => setMeasureTick((t) => t + 1));
    return () => window.cancelAnimationFrame(raf);
  }, [layouts.length]);

  // Recompute on host resize so a window resize or zoom change flows
  // through. Snapshot edits already reach us via the `revisionsKey`
  // dependency above.
  useEffect(() => {
    if (!editorHost) return;
    const ro = new ResizeObserver(() => setMeasureTick((t) => t + 1));
    ro.observe(editorHost);
    const onResize = () => setMeasureTick((t) => t + 1);
    window.addEventListener("resize", onResize);
    return () => {
      ro.disconnect();
      window.removeEventListener("resize", onResize);
    };
  }, [editorHost]);

  if (revisions.length === 0 || columnLeft === null || barLeft === null) return null;

  return (
    <div
      data-testid="tracked-changes-margin"
      className="pointer-events-none absolute inset-0"
      aria-label="Tracked changes margin"
    >
      {layouts.map((l) => (
        <span
          key={`bar-${l.revisionId}`}
          className="pm-revision-changebar"
          aria-hidden
          style={{
            left: barLeft,
            top: l.anchorTop,
            height: l.anchorHeight,
          }}
        />
      ))}
      {layouts.map((l) => (
        <div
          key={`balloon-${l.revisionId}`}
          ref={(el) => {
            if (el) balloonRefs.current.set(l.revisionId, el);
            else balloonRefs.current.delete(l.revisionId);
          }}
          data-testid="tracked-change-balloon"
          data-revision-id={l.revisionId}
          data-revision-type={l.revisionType}
          className="tracked-change-balloon pointer-events-auto"
          style={{
            top: l.balloonTop,
            left: columnLeft,
            width: BALLOON_WIDTH,
          }}
        >
          <div className="text-foreground font-medium">{l.author || "Unknown"}</div>
          <div className="text-secondary">
            {l.revisionType === "del" ? "hat gelöscht: " : "hat eingefügt: "}
            <span className="text-foreground font-medium">
              {snippet(l.previewText, 80) || "(empty)"}
            </span>
          </div>
          <div className="flex items-center gap-1 pt-0.5">
            <button
              type="button"
              title="Accept change"
              aria-label={`Accept change ${l.revisionId}`}
              onClick={() => void onAccept(l.revisionId)}
              className={cn(
                "inline-flex items-center gap-1 rounded px-1.5 py-0.5 text-[10px] font-medium",
                "text-[var(--success)] hover:bg-[var(--success)]/10"
              )}
            >
              <Check size={11} />
              Accept
            </button>
            <button
              type="button"
              title="Reject change"
              aria-label={`Reject change ${l.revisionId}`}
              onClick={() => void onReject(l.revisionId)}
              className={cn(
                "inline-flex items-center gap-1 rounded px-1.5 py-0.5 text-[10px] font-medium",
                "text-[var(--error)] hover:bg-[var(--error)]/10"
              )}
            >
              <X size={11} />
              Reject
            </button>
          </div>
        </div>
      ))}
    </div>
  );
}

function cssEscape(value: string): string {
  if (typeof CSS !== "undefined" && typeof CSS.escape === "function") return CSS.escape(value);
  return value.replace(/[^a-zA-Z0-9_-]/g, (c) => `\\${c}`);
}
