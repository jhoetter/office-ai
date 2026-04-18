"use client";

import { useMemo, type ReactNode } from "react";
import type { DocxSnapshot, SectionProperties } from "@officeai/docx";

/**
 * P3.5 / W20 — read-only page ruler.
 *
 * Renders a horizontal bar above the editor surface showing:
 *
 * - Tick marks every 0.5 inches (or 1 cm in metric locales).
 * - Margin guides drawn from the active section's typed
 *   {@link SectionProperties}: a darker band represents the printable
 *   width (page width − left/right margins), a lighter band the page
 *   margins.
 *
 * Pure render, no PM coupling — the editor passes the snapshot in and
 * the ruler resolves the *trailing implicit section's* geometry to
 * match what the user is most likely seeing on page 1. Per-paragraph
 * resolution lands when the page-aware caret context (P3.6) is wired
 * in.
 *
 * P4 / R8 makes the margin handles draggable; for now the ruler is
 * purely informational so the user can see Word's "you have a 1-inch
 * margin" cue without leaving the app.
 */
export interface PageRulerProps {
  snapshot: DocxSnapshot | null;
  /** Optional zoom factor (1.0 = 100 %). Defaults to 1.0. */
  zoom?: number;
}

const TWIPS_PER_INCH = 1440;
const TWIPS_PER_CM = 567; // 1440 / 2.54

const DEFAULT_PAGE_WIDTH_TWIPS = 12240; // US-letter
const DEFAULT_MARGIN_TWIPS = 1440;

export function PageRuler(props: PageRulerProps): ReactNode {
  const { snapshot, zoom = 1 } = props;

  const geometry = useMemo(() => resolveSectionGeometry(snapshot), [snapshot]);
  const useMetric = useMemo(() => isMetricLocale(), []);

  const totalUnits = useMetric ? geometry.pageWidthTwips / TWIPS_PER_CM : geometry.pageWidthTwips / TWIPS_PER_INCH;
  const leftMarginUnits = useMetric ? geometry.leftMarginTwips / TWIPS_PER_CM : geometry.leftMarginTwips / TWIPS_PER_INCH;
  const rightMarginUnits = useMetric ? geometry.rightMarginTwips / TWIPS_PER_CM : geometry.rightMarginTwips / TWIPS_PER_INCH;

  const ticks: number[] = [];
  const tickStep = useMetric ? 1 : 0.5;
  for (let v = 0; v <= totalUnits + 0.001; v += tickStep) {
    ticks.push(v);
  }

  const unitLabel = useMetric ? "cm" : "in";

  return (
    <div
      className="docx-page-ruler relative h-5 w-full select-none border-b border-divider bg-surface text-[9px] text-secondary"
      style={{ transform: `scaleX(${zoom})`, transformOrigin: "top left" }}
      role="presentation"
      data-testid="page-ruler"
      data-units={unitLabel}
    >
      {/* Margin shading — left margin */}
      <div
        className="absolute inset-y-0 left-0 bg-divider/40"
        style={{ width: `${(leftMarginUnits / totalUnits) * 100}%` }}
        aria-hidden
      />
      {/* Margin shading — right margin */}
      <div
        className="absolute inset-y-0 right-0 bg-divider/40"
        style={{ width: `${(rightMarginUnits / totalUnits) * 100}%` }}
        aria-hidden
      />
      {/* Tick marks */}
      {ticks.map((u) => {
        const pct = (u / totalUnits) * 100;
        const isMajor = useMetric ? Math.round(u) === u : Math.round(u) === u;
        return (
          <div
            key={u}
            className="absolute top-0 flex h-full flex-col items-center"
            style={{ left: `${pct}%`, transform: "translateX(-50%)" }}
          >
            <div
              className="bg-secondary"
              style={{ width: 1, height: isMajor ? "70%" : "40%" }}
            />
            {isMajor && u !== 0 && (
              <span className="mt-[1px] tabular-nums">{Math.round(u)}</span>
            )}
          </div>
        );
      })}
    </div>
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

function isMetricLocale(): boolean {
  if (typeof navigator === "undefined") return false;
  const lang = (navigator.language || "en-US").toLowerCase();
  // US, UK, Liberia, Myanmar use imperial. Everyone else: metric.
  if (lang.startsWith("en-us")) return false;
  if (lang.startsWith("en-gb")) return false;
  if (lang.startsWith("en-lr")) return false;
  if (lang === "my-mm") return false;
  return true;
}
