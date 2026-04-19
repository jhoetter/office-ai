"use client";

import type { CSSProperties, ReactNode } from "react";
import { colToLetter } from "@officeai/xlsx";

/**
 * Static, zero-state stand-in for {@link Grid} that paints the
 * spreadsheet chrome (corner + column letters + row numbers + a
 * tiled gridline background) while the real `XlsxAgent` is still
 * being constructed from the sample bytes.
 *
 * Mounting a real grid requires `buildSampleXlsx()` (~50 ms of
 * JSZip work) followed by `XlsxAgent.fromBuffer()` (parse +
 * snapshot build, easily another 100 ms cold). For that brief
 * window the editor used to render the "Open a workbook" empty
 * state, which made the user feel like they had to do something
 * before the app was usable. Showing the bare frame instead means
 * the surface looks alive on the very first paint — same column
 * letters in the same positions as the live grid, and selection /
 * editing affordances click in as soon as the agent commits.
 *
 * Cheap on purpose: no virtualization, no event handlers, no
 * effects, no measurement. Body cells aren't real DOM nodes —
 * they're a pair of repeating linear-gradients that draw the
 * grid lines for free regardless of viewport size. Total cost is
 * a few dozen header divs.
 */
export function GridSkeleton(): ReactNode {
  return (
    <div
      data-testid="xlsx-grid-skeleton"
      aria-hidden="true"
      style={{
        position: "relative",
        width: "100%",
        height: "100%",
        overflow: "hidden",
        background: "var(--background)",
        border: "1px solid var(--divider)",
        borderRadius: 6,
      }}
    >
      <div style={CORNER_STYLE} />
      <div style={COL_BAND_STYLE}>
        {COL_HEADER_LABELS.map((label, i) => (
          <div
            key={label}
            style={{
              position: "absolute",
              top: 0,
              left: HEADER_COL_WIDTH + i * COL_WIDTH,
              width: COL_WIDTH,
              height: HEADER_ROW_HEIGHT,
              borderRight: "1px solid var(--divider)",
              borderBottom: "1px solid var(--divider)",
              background: "var(--surface)",
              color: "var(--secondary)",
              fontSize: 11,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
            }}
          >
            {label}
          </div>
        ))}
      </div>
      <div style={ROW_BAND_STYLE}>
        {ROW_HEADER_LABELS.map((label, i) => (
          <div
            key={label}
            style={{
              position: "absolute",
              top: HEADER_ROW_HEIGHT + i * ROW_HEIGHT,
              left: 0,
              width: HEADER_COL_WIDTH,
              height: ROW_HEIGHT,
              borderRight: "1px solid var(--divider)",
              borderBottom: "1px solid var(--divider)",
              background: "var(--surface)",
              color: "var(--secondary)",
              fontSize: 11,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
            }}
          >
            {label}
          </div>
        ))}
      </div>
      <div style={BODY_STYLE} />
    </div>
  );
}

// Must mirror Grid.tsx — keep them in sync if those constants ever
// move into a shared module so the skeleton frame stays pixel-identical
// to the real grid as it loads in over the top.
const ROW_HEIGHT = 24;
const COL_WIDTH = 88;
const HEADER_ROW_HEIGHT = 24;
const HEADER_COL_WIDTH = 48;

// Generous enough to cover a 4K viewport without measuring it.
// Off-screen labels are cheap; the alternative (measure → paint)
// would defeat the point of an instant skeleton.
const VISIBLE_COLS = 32;
const VISIBLE_ROWS = 80;

const COL_HEADER_LABELS = Array.from({ length: VISIBLE_COLS }, (_, i) => colToLetter(i));
const ROW_HEADER_LABELS = Array.from({ length: VISIBLE_ROWS }, (_, i) => String(i + 1));

const CORNER_STYLE: CSSProperties = {
  position: "absolute",
  top: 0,
  left: 0,
  width: HEADER_COL_WIDTH,
  height: HEADER_ROW_HEIGHT,
  background: "var(--surface)",
  borderRight: "1px solid var(--divider)",
  borderBottom: "1px solid var(--divider)",
  zIndex: 3,
};

const COL_BAND_STYLE: CSSProperties = {
  position: "absolute",
  top: 0,
  left: 0,
  width: 0,
  height: 0,
  zIndex: 2,
};

const ROW_BAND_STYLE: CSSProperties = {
  position: "absolute",
  top: 0,
  left: 0,
  width: 0,
  height: 0,
  zIndex: 2,
};

// The body is purely decorative — gridlines drawn by tiled
// linear-gradients so the cost stays O(1) regardless of how many
// cells fit in the viewport.
const BODY_STYLE: CSSProperties = {
  position: "absolute",
  top: HEADER_ROW_HEIGHT,
  left: HEADER_COL_WIDTH,
  right: 0,
  bottom: 0,
  backgroundImage: [
    "linear-gradient(to right, var(--divider) 1px, transparent 1px)",
    "linear-gradient(to bottom, var(--divider) 1px, transparent 1px)",
  ].join(", "),
  backgroundSize: `${COL_WIDTH}px ${ROW_HEIGHT}px, ${COL_WIDTH}px ${ROW_HEIGHT}px`,
  backgroundPosition: "0 0, 0 0",
  opacity: 0.6,
};
