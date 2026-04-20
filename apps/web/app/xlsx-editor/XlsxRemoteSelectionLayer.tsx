"use client";

import type { ReactNode } from "react";
import type { AwarenessState } from "@officeai/realtime";
import { colors } from "@officeai/design-tokens";
import type { AxisLookup } from "./gridDimensions";
import { parseRangeA1 } from "./parseRangeA1";

/**
 * Remote-selection overlay for the XLSX grid.
 *
 * Each remote peer publishes an `XlsxSelection` (`sheetName`,
 * `anchor`, `range`) via the realtime awareness layer. For every
 * peer whose `sheetName` matches the active sheet we draw a 2px
 * colored outline around their selected range, with a small name
 * tag pinned to the top-right corner. Outlines are painted in the
 * same scrolling overlay layer that hosts the local marquee /
 * marching ants / image / chart overlays so they pan with the
 * grid for free.
 *
 * Peers on a different sheet are silently filtered (see the sheet
 * tab dots in `SheetTabBar` for cross-sheet visibility).
 */

export interface RemoteXlsxPeer {
  readonly clientId: number;
  readonly state: AwarenessState;
}

export interface XlsxRemoteSelectionLayerProps {
  readonly peers: ReadonlyArray<RemoteXlsxPeer>;
  readonly activeSheetName: string;
  readonly colXs: AxisLookup;
  readonly rowYs: AxisLookup;
  readonly headerOffset: { readonly x: number; readonly y: number };
}

interface RemoteRect {
  readonly clientId: number;
  readonly name: string;
  readonly color: string;
  readonly top: number;
  readonly left: number;
  readonly width: number;
  readonly height: number;
}

export function XlsxRemoteSelectionLayer(props: XlsxRemoteSelectionLayerProps): ReactNode {
  const { peers, activeSheetName, colXs, rowYs, headerOffset } = props;
  const rects: RemoteRect[] = [];
  for (const peer of peers) {
    const c = peer.state.cursor;
    if (!c || c.product !== "xlsx") continue;
    if (c.sheetName !== activeSheetName) continue;
    const parsed = parseRangeA1(c.range);
    if (!parsed) continue;
    const { r1, c1, r2, c2 } = parsed;
    if (r1 >= rowYs.length - 1 || c1 >= colXs.length - 1) continue;
    const r1Clamped = Math.max(0, Math.min(rowYs.length - 2, r1));
    const r2Clamped = Math.max(0, Math.min(rowYs.length - 2, r2));
    const c1Clamped = Math.max(0, Math.min(colXs.length - 2, c1));
    const c2Clamped = Math.max(0, Math.min(colXs.length - 2, c2));
    const top = headerOffset.y + (rowYs[r1Clamped] ?? 0);
    const left = headerOffset.x + (colXs[c1Clamped] ?? 0);
    const width = (colXs[c2Clamped + 1] ?? 0) - (colXs[c1Clamped] ?? 0);
    const height = (rowYs[r2Clamped + 1] ?? 0) - (rowYs[r1Clamped] ?? 0);
    rects.push({
      clientId: peer.clientId,
      name: peer.state.user.name,
      color: peer.state.user.color,
      top,
      left,
      width: Math.max(2, width),
      height: Math.max(2, height),
    });
  }
  if (rects.length === 0) return null;
  return (
    <>
      {rects.map((r) => (
        <div
          key={`xlsx-remote-${r.clientId}`}
          data-testid="xlsx-remote-selection"
          data-peer-color={r.color}
          aria-hidden
          style={{
            position: "absolute",
            top: r.top,
            left: r.left,
            width: r.width,
            height: r.height,
            border: `2px solid ${r.color}`,
            boxSizing: "border-box",
            pointerEvents: "none",
            zIndex: 4,
          }}
        >
          <div
            style={{
              position: "absolute",
              top: -16,
              left: -1,
              padding: "1px 4px",
              fontSize: 10,
              fontWeight: 500,
              color: colors.background,
              backgroundColor: r.color,
              borderRadius: "2px 2px 2px 0",
              whiteSpace: "nowrap",
            }}
          >
            {r.name}
          </div>
        </div>
      ))}
    </>
  );
}
