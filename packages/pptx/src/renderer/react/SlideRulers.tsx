import * as React from "react";
import type { SlideSize } from "../../model/types.js";
import { EMU_PER_INCH } from "../layout/units.js";

/**
 * EMU per centimetre. PowerPoint stores 1 cm = 360000 EMU (since
 * 1 in = 914400 EMU and 1 in = 2.54 cm).
 */
const EMU_PER_CM = 360000;

/**
 * PowerPoint-style linear ruler painted around the editor's slide card.
 *
 * Two instances are rendered by {@link SlideCanvas}:
 *
 * - `axis="h"` — a horizontal strip pinned to the top of the stage
 *   that spans the full stage width. Tick `0` is anchored to the
 *   slide card's left edge; positive ticks run right across the
 *   slide, and the strip extends into the surrounding scratch margin
 *   on both sides (negative ticks on the left, ticks beyond
 *   `slideWidthInUnits` on the right).
 * - `axis="v"` — a vertical strip pinned to the left of the stage.
 *   Same convention, rotated 90°.
 *
 * The component is locale-agnostic: the host editor picks the unit
 * (cm in metric locales, in elsewhere — see `apps/web/app/lib/ruler/units.ts`)
 * and passes it in. This keeps the pptx package free of `navigator`
 * lookups and trivially testable.
 *
 * Pure presentational; no input handling. Selection-aware highlight
 * (the darker band PowerPoint paints over the active shape's bounds)
 * is intentionally out of scope for the v1 — the rulers' job is to
 * show the slide scale, not to mirror selection state.
 */
export interface SlideRulersProps {
  readonly axis: "h" | "v";
  readonly slideSize: SlideSize;
  /**
   * Pixel rect of the slide card inside the parent stage (the same
   * `slidePx*` values the canvas already computes via
   * `useStageLayout`). Origin (tick 0) is anchored to `slidePxLeft`
   * for the horizontal ruler and `slidePxTop` for the vertical one.
   */
  readonly stage: {
    readonly stageW: number;
    readonly stageH: number;
    readonly slidePxLeft: number;
    readonly slidePxTop: number;
    readonly slidePxW: number;
    readonly slidePxH: number;
  };
  /** Display unit; the host editor picks this from the user's locale. */
  readonly unit: "in" | "cm";
}

/** Width of the ruler strip in CSS pixels. Matches the docx PageRuler. */
export const RULER_THICKNESS_PX = 20;

export function SlideRulers(props: SlideRulersProps): React.ReactElement {
  const { axis, slideSize, stage, unit } = props;
  const horizontal = axis === "h";

  const emuPerUnit = unit === "cm" ? EMU_PER_CM : EMU_PER_INCH;
  const step = unit === "cm" ? 1 : 0.5;

  const slidePxOrigin = horizontal ? stage.slidePxLeft : stage.slidePxTop;
  const slidePxLength = horizontal ? stage.slidePxW : stage.slidePxH;
  const stageLength = horizontal ? stage.stageW : stage.stageH;
  const slideLengthEmu = horizontal ? slideSize.cxEmu : slideSize.cyEmu;
  const slideLengthUnits = slideLengthEmu / emuPerUnit;

  // Pixels per unit, derived from the rendered slide size so the ruler
  // tracks zoom changes for free (the parent re-measures via the
  // stage layout's ResizeObserver).
  const pxPerUnit = slidePxLength > 0 ? slidePxLength / slideLengthUnits : 0;

  // First/last tick covers the entire stage strip, anchored to slide 0.
  const firstUnit = pxPerUnit > 0 ? -slidePxOrigin / pxPerUnit : 0;
  const lastUnit = pxPerUnit > 0 ? (stageLength - slidePxOrigin) / pxPerUnit : 0;
  const firstK = Math.ceil(firstUnit / step - 1e-6);
  const lastK = Math.floor(lastUnit / step + 1e-6);

  const ticks: { readonly value: number; readonly major: boolean }[] = [];
  for (let k = firstK; k <= lastK; k++) {
    const value = k * step;
    ticks.push({ value, major: Math.abs(value - Math.round(value)) < 1e-6 });
  }

  const containerStyle: React.CSSProperties = horizontal
    ? {
        position: "absolute",
        top: 0,
        left: 0,
        width: "100%",
        height: `${RULER_THICKNESS_PX}px`,
      }
    : {
        position: "absolute",
        top: 0,
        left: 0,
        width: `${RULER_THICKNESS_PX}px`,
        height: "100%",
      };

  // The slide-bounds band gets a slightly lighter background so the
  // user can see at a glance how far across the slide they are. Mirrors
  // the docx PageRuler's "darker bands at the margins" affordance,
  // inverted because in pptx everything outside the slide is scratch.
  const bandStyle: React.CSSProperties = horizontal
    ? {
        position: "absolute",
        top: 0,
        bottom: 0,
        left: `${slidePxOrigin}px`,
        width: `${slidePxLength}px`,
      }
    : {
        position: "absolute",
        left: 0,
        right: 0,
        top: `${slidePxOrigin}px`,
        height: `${slidePxLength}px`,
      };

  return (
    <div
      data-testid={horizontal ? "pptx-ruler-h" : "pptx-ruler-v"}
      data-units={unit}
      style={{
        ...containerStyle,
        pointerEvents: "none",
        userSelect: "none",
        background: "var(--surface, #ffffff)",
        borderRight: horizontal ? undefined : "1px solid var(--divider, #e5e5e5)",
        borderBottom: horizontal ? "1px solid var(--divider, #e5e5e5)" : undefined,
        color: "var(--secondary, #6b7280)",
        fontSize: 9,
        zIndex: 5,
        overflow: "hidden",
      }}
      role="presentation"
    >
      {pxPerUnit > 0 ? (
        <div
          aria-hidden
          style={{
            ...bandStyle,
            background: "var(--surface, #ffffff)",
            // Slightly brighter than the scratch part of the strip.
            filter: "brightness(1.04)",
          }}
        />
      ) : null}
      {ticks.map(({ value, major }) => {
        const px = slidePxOrigin + value * pxPerUnit;
        if (horizontal) {
          return (
            <div
              key={`h-${value.toFixed(4)}`}
              style={{
                position: "absolute",
                left: `${px}px`,
                top: 0,
                bottom: 0,
                transform: "translateX(-50%)",
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                justifyContent: "flex-start",
              }}
            >
              <div
                style={{
                  width: 1,
                  height: major ? "70%" : "40%",
                  background: "currentColor",
                  opacity: 0.6,
                }}
              />
              {major && Math.round(value) !== 0 ? (
                <span
                  style={{
                    marginTop: 1,
                    fontVariantNumeric: "tabular-nums",
                  }}
                >
                  {Math.abs(Math.round(value))}
                </span>
              ) : null}
            </div>
          );
        }
        return (
          <div
            key={`v-${value.toFixed(4)}`}
            style={{
              position: "absolute",
              top: `${px}px`,
              left: 0,
              right: 0,
              transform: "translateY(-50%)",
              display: "flex",
              flexDirection: "row",
              alignItems: "center",
              justifyContent: "flex-start",
            }}
          >
            <div
              style={{
                height: 1,
                width: major ? "70%" : "40%",
                background: "currentColor",
                opacity: 0.6,
              }}
            />
            {major && Math.round(value) !== 0 ? (
              <span
                style={{
                  marginLeft: 1,
                  fontVariantNumeric: "tabular-nums",
                  // Rotate -90deg keeps numbers readable from the side
                  // without turning the strip into a column of stacked
                  // single digits — same trick PowerPoint uses.
                  writingMode: "vertical-rl",
                  transform: "rotate(180deg)",
                  lineHeight: 1,
                }}
              >
                {Math.abs(Math.round(value))}
              </span>
            ) : null}
          </div>
        );
      })}
    </div>
  );
}
