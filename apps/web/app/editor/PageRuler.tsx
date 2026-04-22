"use client";

import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import type { DocxSnapshot, SectionProperties } from "@officeai/docx";
import { useTranslator } from "@/lib/i18n";
import {
  TWIPS_PER_CM,
  TWIPS_PER_INCH,
  buildTicks,
  isMajorTick,
  isMetricLocale,
  twipsToUnit,
} from "@/lib/ruler/units";

/**
 * P3.5 / W20 + B3 — page ruler with draggable margin handles.
 *
 * Renders a horizontal bar above the editor surface showing:
 *
 * - Tick marks every 0.5 inches (or 1 cm in metric locales).
 * - Margin guides drawn from the active section's typed
 *   {@link SectionProperties}: a darker band represents the page
 *   margins, a lighter band the printable width (page width −
 *   left/right margins).
 * - Two draggable handles flagged "M". Dragging dispatches
 *   {@link onMarginsChange} with the new left / right margins in
 *   twips so the editor can route the mutation through the
 *   `docx:set-page-setup` command. While dragging, the visual
 *   margins update locally so the user sees an immediate response;
 *   the snapshot becomes the source of truth on pointerup.
 *
 * Drag is single-axis (horizontal) and clamped to keep the printable
 * width >= 360 twips (¼"), matching Word's behaviour.
 */
export interface PageRulerProps {
  snapshot: DocxSnapshot | null;
  /** Optional zoom factor (1.0 = 100 %). Defaults to 1.0. */
  zoom?: number;
  /**
   * Called when the user drops a margin handle. Twips, absolute
   * (not deltas). Omit to render the ruler in read-only mode.
   */
  onMarginsChange?: (next: { left: number; right: number }) => void;
  /** Optional click handler — fires when the ruler body is clicked
   * (i.e. not on a handle). Used to open the Page Setup dialog. */
  onOpenPageSetup?: () => void;
}

const DEFAULT_PAGE_WIDTH_TWIPS = 12240; // US-letter
const DEFAULT_MARGIN_TWIPS = 1440;
const MIN_PRINTABLE_TWIPS = 360; // ¼"

export function PageRuler(props: PageRulerProps): ReactNode {
  const { snapshot, zoom = 1, onMarginsChange, onOpenPageSetup } = props;
  const { t } = useTranslator();

  const baseGeometry = useMemo(() => resolveSectionGeometry(snapshot), [snapshot]);
  const useMetric = useMemo(() => isMetricLocale(), []);
  const containerRef = useRef<HTMLDivElement | null>(null);
  // While dragging we mirror the live margin into local state so the
  // user sees the handle move at 60fps without waiting for the
  // command bus to round-trip through the snapshot.
  const [draftLeft, setDraftLeft] = useState<number | null>(null);
  const [draftRight, setDraftRight] = useState<number | null>(null);
  // Whichever side is currently being dragged (or null when idle).
  // Used to bind the document-level listeners and to suppress the
  // hover cursor on the inert side.
  const dragRef = useRef<"left" | "right" | null>(null);

  const leftMarginTwips = draftLeft ?? baseGeometry.leftMarginTwips;
  const rightMarginTwips = draftRight ?? baseGeometry.rightMarginTwips;
  const pageWidthTwips = baseGeometry.pageWidthTwips;

  const unit = useMetric ? "cm" : "in";
  const totalUnits = twipsToUnit(pageWidthTwips, unit);
  const leftMarginUnits = twipsToUnit(leftMarginTwips, unit);
  const rightMarginUnits = twipsToUnit(rightMarginTwips, unit);

  const tickStep = useMetric ? 1 : 0.5;
  const ticks = buildTicks(0, totalUnits, tickStep);

  const unitLabel = unit;

  const draggable = typeof onMarginsChange === "function";

  const beginDrag = useCallback(
    (side: "left" | "right", startEvent: React.PointerEvent<HTMLDivElement>) => {
      if (!draggable) return;
      const container = containerRef.current;
      if (!container) return;
      startEvent.preventDefault();
      startEvent.stopPropagation();
      dragRef.current = side;
      const containerRect = container.getBoundingClientRect();
      // The ruler is inside a `transform: scaleX(zoom)` wrapper but
      // we measure the rendered rect, which already accounts for
      // the scale. So twips per CSS px is constant regardless of
      // zoom: containerWidth(px) ↔ pageWidth(twips).
      const twipsPerPx = pageWidthTwips / containerRect.width;

      const onMove = (e: PointerEvent) => {
        const xPx = e.clientX - containerRect.left;
        const xTwips = clamp(xPx * twipsPerPx, 0, pageWidthTwips);
        if (side === "left") {
          const maxLeft = pageWidthTwips - rightMarginTwips - MIN_PRINTABLE_TWIPS;
          setDraftLeft(clamp(xTwips, 0, Math.max(0, maxLeft)));
        } else {
          const newRight = pageWidthTwips - xTwips;
          const maxRight = pageWidthTwips - leftMarginTwips - MIN_PRINTABLE_TWIPS;
          setDraftRight(clamp(newRight, 0, Math.max(0, maxRight)));
        }
      };
      const onUp = () => {
        document.removeEventListener("pointermove", onMove);
        document.removeEventListener("pointerup", onUp);
        document.removeEventListener("pointercancel", onUp);
        const finalLeft = side === "left" ? (draftLeftRef.current ?? leftMarginTwips) : leftMarginTwips;
        const finalRight = side === "right" ? (draftRightRef.current ?? rightMarginTwips) : rightMarginTwips;
        dragRef.current = null;
        setDraftLeft(null);
        setDraftRight(null);
        onMarginsChange?.({
          left: Math.round(finalLeft),
          right: Math.round(finalRight),
        });
      };
      document.addEventListener("pointermove", onMove);
      document.addEventListener("pointerup", onUp);
      document.addEventListener("pointercancel", onUp);
    },
    [draggable, leftMarginTwips, onMarginsChange, pageWidthTwips, rightMarginTwips]
  );

  // Mirror draft state into refs so the pointerup handler sees the
  // latest values (the handler is bound once in beginDrag and
  // closures over the original draft state).
  const draftLeftRef = useRef<number | null>(null);
  const draftRightRef = useRef<number | null>(null);
  useEffect(() => {
    draftLeftRef.current = draftLeft;
  }, [draftLeft]);
  useEffect(() => {
    draftRightRef.current = draftRight;
  }, [draftRight]);

  return (
    <div
      ref={containerRef}
      className="docx-page-ruler relative h-5 w-full select-none border-b border-divider bg-surface text-[9px] text-secondary"
      style={{ transform: `scaleX(${zoom})`, transformOrigin: "top left" }}
      role="presentation"
      data-testid="page-ruler"
      data-units={unitLabel}
      onDoubleClick={(e) => {
        if (!onOpenPageSetup) return;
        // Double-click on the ruler body opens Page Setup, mirroring
        // Word's affordance. We intentionally use dblclick (not click)
        // so a stray click while dragging never accidentally pops the
        // dialog.
        e.preventDefault();
        onOpenPageSetup();
      }}
      title={onOpenPageSetup ? t("docx.pageRuler.openPageSetup") : undefined}
    >
      <div
        className="absolute inset-y-0 left-0 bg-divider/40"
        style={{ width: `${(leftMarginUnits / totalUnits) * 100}%` }}
        aria-hidden
      />
      <div
        className="absolute inset-y-0 right-0 bg-divider/40"
        style={{ width: `${(rightMarginUnits / totalUnits) * 100}%` }}
        aria-hidden
      />
      {ticks.map((u) => {
        const pct = (u / totalUnits) * 100;
        const major = isMajorTick(u);
        return (
          <div
            key={u}
            className="absolute top-0 flex h-full flex-col items-center"
            style={{ left: `${pct}%`, transform: "translateX(-50%)" }}
          >
            <div className="bg-secondary" style={{ width: 1, height: major ? "70%" : "40%" }} />
            {major && u !== 0 && <span className="mt-[1px] tabular-nums">{Math.round(u)}</span>}
          </div>
        );
      })}
      {draggable ? (
        <>
          <MarginHandle
            side="left"
            positionPct={(leftMarginUnits / totalUnits) * 100}
            twips={leftMarginTwips}
            useMetric={useMetric}
            onPointerDown={(e) => beginDrag("left", e)}
            active={dragRef.current === "left"}
            t={t}
          />
          <MarginHandle
            side="right"
            positionPct={100 - (rightMarginUnits / totalUnits) * 100}
            twips={rightMarginTwips}
            useMetric={useMetric}
            onPointerDown={(e) => beginDrag("right", e)}
            active={dragRef.current === "right"}
            t={t}
          />
        </>
      ) : null}
    </div>
  );
}

interface MarginHandleProps {
  side: "left" | "right";
  positionPct: number;
  twips: number;
  useMetric: boolean;
  active: boolean;
  onPointerDown: (e: React.PointerEvent<HTMLDivElement>) => void;
  t: (key: string, vars?: Readonly<Record<string, string | number>>) => string;
}

function MarginHandle(props: MarginHandleProps) {
  const { side, positionPct, twips, useMetric, active, onPointerDown, t } = props;
  const factor = useMetric ? TWIPS_PER_CM : TWIPS_PER_INCH;
  const labelValue = (twips / factor).toFixed(2);
  const unit = useMetric ? "cm" : "in";
  const ariaKey = side === "left" ? "docx.pageRuler.leftMargin" : "docx.pageRuler.rightMargin";
  const titleKey =
    side === "left" ? "docx.pageRuler.leftMarginValue" : "docx.pageRuler.rightMarginValue";
  return (
    <div
      role="slider"
      aria-label={t(ariaKey)}
      aria-valuenow={Number(labelValue)}
      aria-valuemin={0}
      data-testid={`page-ruler-handle-${side}`}
      onPointerDown={onPointerDown}
      className={`absolute top-0 h-full w-2 -translate-x-1/2 cursor-ew-resize transition-colors ${
        active ? "bg-accent" : "bg-secondary/50 hover:bg-accent/70"
      }`}
      style={{ left: `${positionPct}%` }}
      title={t(titleKey, { value: labelValue, unit })}
    />
  );
}

interface ResolvedGeometry {
  readonly pageWidthTwips: number;
  readonly leftMarginTwips: number;
  readonly rightMarginTwips: number;
}

function resolveSectionGeometry(snapshot: DocxSnapshot | null): ResolvedGeometry {
  const props = trailingSectionProperties(snapshot);
  const pageWidth = props?.pgSz?.w ?? DEFAULT_PAGE_WIDTH_TWIPS;
  const left = props?.pgMar?.left ?? DEFAULT_MARGIN_TWIPS;
  const right = props?.pgMar?.right ?? DEFAULT_MARGIN_TWIPS;
  return { pageWidthTwips: pageWidth, leftMarginTwips: left, rightMarginTwips: right };
}

function trailingSectionProperties(snapshot: DocxSnapshot | null): SectionProperties | null {
  if (!snapshot) return null;
  const body = snapshot.root.body;
  for (let i = body.length - 1; i >= 0; i--) {
    const block = body[i];
    if (block.kind === "section-break") return block.properties;
  }
  return null;
}

function clamp(n: number, lo: number, hi: number): number {
  if (!Number.isFinite(n)) return lo;
  if (n < lo) return lo;
  if (n > hi) return hi;
  return n;
}
