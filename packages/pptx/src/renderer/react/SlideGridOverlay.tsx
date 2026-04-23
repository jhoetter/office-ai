import * as React from "react";
import type { SlideSize } from "../../model/types.js";
import { EMU_PER_INCH, EMU_PER_PX_AT_96DPI } from "../layout/units.js";

const EMU_PER_CM = 360000;

/**
 * PowerPoint-style gridlines overlay painted directly on the slide
 * card. Toggled by the View ribbon's "Gridlines" button.
 *
 * Visual: a faint major grid at every `0.5 in` / `1 cm` interval
 * (matching the ruler's tick step) plus an even fainter minor grid at
 * 1/4 of the major step. Lines are sub-pixel-thick at 96 DPI. The
 * overlay is sized to the slide rectangle and clipped to it so
 * off-slide content in the scratch area stays free of grid noise —
 * PowerPoint behaves the same way.
 *
 * Pure presentational and locale-agnostic; the host editor passes a
 * `unit` prop derived from the user's locale.
 *
 * Two flavours are exported:
 *
 * - {@link SlideGridOverlay} renders inside an existing React/SVG tree
 *   (currently unused, but useful for tests and for hosts that don't
 *   share the dangerously-set-inner-HTML strategy).
 * - {@link slideGridSvgString} returns the same markup as a string,
 *   for callers (specifically `SlideCanvas`) that build their slide
 *   SVG via `innerHTML` so shapes naturally paint over the grid.
 */
export interface SlideGridOverlayProps {
  readonly slideSize: SlideSize;
  readonly unit: "in" | "cm";
}

interface GridGeometry {
  readonly slideW: number;
  readonly slideH: number;
  readonly major: number;
  readonly minor: number;
}

function gridGeometry(slideSize: SlideSize, unit: "in" | "cm"): GridGeometry {
  const emuPerUnit = unit === "cm" ? EMU_PER_CM : EMU_PER_INCH;
  const majorEmu = unit === "cm" ? emuPerUnit : emuPerUnit / 2; // 1 cm or 0.5 in
  const minorEmu = majorEmu / 4;
  return {
    slideW: slideSize.cxEmu / EMU_PER_PX_AT_96DPI,
    slideH: slideSize.cyEmu / EMU_PER_PX_AT_96DPI,
    major: majorEmu / EMU_PER_PX_AT_96DPI,
    minor: minorEmu / EMU_PER_PX_AT_96DPI,
  };
}

const MINOR_PATTERN_ID = "pptx-grid-minor";
const MAJOR_PATTERN_ID = "pptx-grid-major";

export function SlideGridOverlay(props: SlideGridOverlayProps): React.ReactElement {
  const { slideSize, unit } = props;
  const { slideW, slideH, major, minor } = gridGeometry(slideSize, unit);
  return (
    <g data-testid="pptx-grid-overlay" pointerEvents="none">
      <defs>
        <pattern id={MINOR_PATTERN_ID} x={0} y={0} width={minor} height={minor} patternUnits="userSpaceOnUse">
          <path
            d={`M ${minor} 0 L 0 0 0 ${minor}`}
            fill="none"
            stroke="#64748b"
            strokeWidth={0.5}
            opacity={0.18}
          />
        </pattern>
        <pattern id={MAJOR_PATTERN_ID} x={0} y={0} width={major} height={major} patternUnits="userSpaceOnUse">
          <rect width={major} height={major} fill={`url(#${MINOR_PATTERN_ID})`} />
          <path
            d={`M ${major} 0 L 0 0 0 ${major}`}
            fill="none"
            stroke="#64748b"
            strokeWidth={0.75}
            opacity={0.45}
          />
        </pattern>
      </defs>
      <rect x={0} y={0} width={slideW} height={slideH} fill={`url(#${MAJOR_PATTERN_ID})`} />
    </g>
  );
}

/**
 * Stringified twin of {@link SlideGridOverlay} for hosts that build
 * their slide SVG via `dangerouslySetInnerHTML` (notably
 * `SlideCanvas`). Inject this between the slide background `<rect>`
 * and the shape group so the grid sits above the background fill but
 * below every shape — matching PowerPoint's stacking.
 */
export function slideGridSvgString(slideSize: SlideSize, unit: "in" | "cm"): string {
  const { slideW, slideH, major, minor } = gridGeometry(slideSize, unit);
  return [
    `<g data-testid="pptx-grid-overlay" pointer-events="none">`,
    `<defs>`,
    `<pattern id="${MINOR_PATTERN_ID}" x="0" y="0" width="${minor}" height="${minor}" patternUnits="userSpaceOnUse">`,
    `<path d="M ${minor} 0 L 0 0 0 ${minor}" fill="none" stroke="#64748b" stroke-width="0.5" opacity="0.18"/>`,
    `</pattern>`,
    `<pattern id="${MAJOR_PATTERN_ID}" x="0" y="0" width="${major}" height="${major}" patternUnits="userSpaceOnUse">`,
    `<rect width="${major}" height="${major}" fill="url(#${MINOR_PATTERN_ID})"/>`,
    `<path d="M ${major} 0 L 0 0 0 ${major}" fill="none" stroke="#64748b" stroke-width="0.75" opacity="0.45"/>`,
    `</pattern>`,
    `</defs>`,
    `<rect x="0" y="0" width="${slideW}" height="${slideH}" fill="url(#${MAJOR_PATTERN_ID})"/>`,
    `</g>`,
  ].join("");
}
