"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { CSSProperties, ReactNode } from "react";
import { cellKey, colToLetter, type Sheet, type StyleTable } from "@officeai/xlsx";
import {
  containsCell,
  isSingle,
  normalizeSelection,
  singleSelection,
  type CellPos,
  type Selection,
} from "./selection";
import { formatCellValue, styleForCell } from "./styles";

/**
 * Fixed cell geometry. Excel's defaults are 64px wide × 20px tall;
 * we widen a touch so column letters and short labels read better.
 */
const ROW_HEIGHT = 24;
const COL_WIDTH = 88;
const HEADER_ROW_HEIGHT = 24;
const HEADER_COL_WIDTH = 48;

/**
 * Total grid extents. The visible viewport is much smaller (we render
 * only what fits + a small overscan), but the inner spacer must be the
 * full virtual size so the browser scrollbar matches a "real" sheet.
 */
const TOTAL_ROWS = 1000;
const TOTAL_COLS = 26;
const OVERSCAN = 4;

export type GridSelection = Selection;

export interface GridProps {
  readonly sheet: Sheet;
  /** Workbook style table — flattened per-cell to render fonts/fills/etc. */
  readonly styles: StyleTable;
  readonly selection: Selection | null;
  /**
   * Drives selection moves. `extend` mirrors Shift-click / drag-extend:
   * keep `anchor`, replace `focus`. Single-cell mousedown leaves both
   * pinned to the same `pos`.
   */
  readonly onSelect: (pos: CellPos, opts?: { extend?: boolean }) => void;
  /**
   * Called when the user commits an in-cell edit (double-click → type
   * → Enter). Parent dispatches `xlsx:set-cell-value` /
   * `xlsx:set-cell-formula`.
   */
  readonly onCommitEdit: (pos: CellPos, value: string) => void;
}

/**
 * Virtualized grid. The DOM strategy:
 *   - One scrollable container (`position: relative`).
 *   - One full-size inner spacer div so the native scrollbar sizes match
 *     the virtual sheet (`TOTAL_COLS × COL_WIDTH` × `TOTAL_ROWS × ROW_HEIGHT`).
 *   - Body cells: `position: absolute`, only the visible range +
 *     overscan are rendered.
 *   - Column / row / corner headers: `position: absolute`, but their
 *     `top` / `left` is set to `scrollTop` / `scrollLeft` so they track
 *     the viewport edge as the user scrolls (poor man's sticky that
 *     plays nicely with absolutely positioned siblings).
 *   - Active selection draws a single bounding-box marquee absolutely
 *     positioned over the cells (`pointer-events: none`) so it reads
 *     like Excel even when the user drags up/left.
 */
export function Grid(props: GridProps): ReactNode {
  const { sheet, styles, selection, onSelect, onCommitEdit } = props;

  const scrollRef = useRef<HTMLDivElement | null>(null);
  const [scroll, setScroll] = useState({ top: 0, left: 0 });
  const [editing, setEditing] = useState<{ row: number; col: number; draft: string } | null>(null);
  const [viewport, setViewport] = useState({ width: 0, height: 0 });
  // True while the user holds the primary mouse button after a
  // mousedown on a body cell — drag-extends the selection.
  const draggingRef = useRef(false);

  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const onResize = () => setViewport({ width: el.clientWidth, height: el.clientHeight });
    onResize();
    const ro = new ResizeObserver(onResize);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // Stop the drag on any global mouseup, even if the user releases
  // outside the grid (e.g. over the formula bar).
  useEffect(() => {
    const onUp = () => {
      draggingRef.current = false;
    };
    window.addEventListener("mouseup", onUp);
    return () => window.removeEventListener("mouseup", onUp);
  }, []);

  const onScroll = useCallback((ev: React.UIEvent<HTMLDivElement>) => {
    const el = ev.currentTarget;
    setScroll({ top: el.scrollTop, left: el.scrollLeft });
  }, []);

  const visibleH = Math.max(viewport.height - HEADER_ROW_HEIGHT, ROW_HEIGHT);
  const visibleW = Math.max(viewport.width - HEADER_COL_WIDTH, COL_WIDTH);

  const startRow = Math.max(0, Math.floor(scroll.top / ROW_HEIGHT) - OVERSCAN);
  const endRow = Math.min(TOTAL_ROWS - 1, Math.ceil((scroll.top + visibleH) / ROW_HEIGHT) + OVERSCAN);
  const startCol = Math.max(0, Math.floor(scroll.left / COL_WIDTH) - OVERSCAN);
  const endCol = Math.min(TOTAL_COLS - 1, Math.ceil((scroll.left + visibleW) / COL_WIDTH) + OVERSCAN);

  const innerStyle: CSSProperties = {
    position: "relative",
    width: HEADER_COL_WIDTH + TOTAL_COLS * COL_WIDTH,
    height: HEADER_ROW_HEIGHT + TOTAL_ROWS * ROW_HEIGHT,
  };

  // Index merged regions by top-left and by "covered" cells. The
  // top-left key yields the merge so we can render an oversized cell
  // spanning the rectangle; the covered set lets us skip those cells
  // in the per-cell loop so the merged surface looks contiguous.
  const mergeIndex = useMemo(() => {
    const topLeft = new Map<string, { r1: number; c1: number; r2: number; c2: number }>();
    const covered = new Set<string>();
    for (const m of sheet.merges) {
      topLeft.set(cellKey(m.r1, m.c1), m);
      for (let r = m.r1; r <= m.r2; r++) {
        for (let c = m.c1; c <= m.c2; c++) {
          if (r === m.r1 && c === m.c1) continue;
          covered.add(cellKey(r, c));
        }
      }
    }
    return { topLeft, covered };
  }, [sheet.merges]);

  const cellList: ReactNode[] = useMemo(() => {
    const out: ReactNode[] = [];
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const k = cellKey(r, c);
        if (mergeIndex.covered.has(k)) continue;
        const merge = mergeIndex.topLeft.get(k);
        const widthCells = merge ? merge.c2 - merge.c1 + 1 : 1;
        const heightCells = merge ? merge.r2 - merge.r1 + 1 : 1;
        const cell = sheet.cells.get(k);
        const cellStyle = styleForCell(styles, cell?.styleId);
        const display = cell ? formatCellValue(cell.value, cellStyle.effective.numFmtId) : "";
        const inSel = !!selection && containsCell(selection, r, c);
        const anchorCell =
          !!selection && selection.anchor.row === r && selection.anchor.col === c;
        const isEditing = editing?.row === r && editing?.col === c;
        out.push(
          <div
            key={`c-${r}-${c}`}
            data-testid={`cell-${colToLetter(c)}${r + 1}`}
            role="gridcell"
            aria-selected={inSel || undefined}
            onMouseDown={(e) => {
              if (isEditing) return;
              // Suppress the browser's default focus shuffle on
              // mousedown — the parent owns focus management. Without
              // this, clicking a body cell while the formula bar is in
              // point mode steals focus away from the input and the
              // click-to-insert-ref handler never sees `formulaEditing`.
              e.preventDefault();
              draggingRef.current = true;
              onSelect({ row: r, col: c }, { extend: e.shiftKey });
            }}
            onMouseEnter={(e) => {
              if (!draggingRef.current) return;
              // Only "buttons & 1" (primary still pressed) qualifies as drag.
              if ((e.buttons & 1) === 0) {
                draggingRef.current = false;
                return;
              }
              onSelect({ row: r, col: c }, { extend: true });
            }}
            onDoubleClick={() =>
              setEditing({
                row: r,
                col: c,
                draft: cell?.formula ? `=${cell.formula.text}` : display,
              })
            }
            style={{
              position: "absolute",
              top: HEADER_ROW_HEIGHT + r * ROW_HEIGHT,
              left: HEADER_COL_WIDTH + c * COL_WIDTH,
              width: COL_WIDTH * widthCells,
              height: ROW_HEIGHT * heightCells,
              boxSizing: "border-box",
              borderRight: "1px solid var(--divider)",
              borderBottom: "1px solid var(--divider)",
              padding: "0 6px",
              display: "flex",
              alignItems: "center",
              fontSize: 12,
              lineHeight: `${ROW_HEIGHT - 2}px`,
              // The bounding-box marquee draws the outline; per-cell
              // background only differentiates the anchor cell from the
              // rest of the selection (Excel-like white anchor).
              background: inSel
                ? anchorCell
                  ? "var(--background)"
                  : "var(--ai-violet-light)"
                : "var(--background)",
              color: "var(--foreground)",
              overflow: "hidden",
              whiteSpace: "nowrap",
              textOverflow: "ellipsis",
              cursor: "cell",
              userSelect: "none",
              zIndex: anchorCell ? 1 : 0,
              ...cellStyle.css,
              // Selection background wins over a per-cell fill so the
              // user can still see what's highlighted, except for the
              // anchor where Excel keeps the cell's real fill.
              ...(inSel && !anchorCell
                ? { background: "var(--ai-violet-light)" }
                : {}),
            }}
          >
            {isEditing ? (
              <input
                autoFocus
                value={editing.draft}
                onChange={(e) => setEditing({ row: r, col: c, draft: e.target.value })}
                onBlur={() => setEditing(null)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    e.preventDefault();
                    onCommitEdit({ row: r, col: c }, editing.draft);
                    setEditing(null);
                  } else if (e.key === "Escape") {
                    e.preventDefault();
                    setEditing(null);
                  }
                }}
                style={{
                  width: "100%",
                  height: "100%",
                  border: "none",
                  outline: "none",
                  background: "var(--background)",
                  font: "inherit",
                  color: "inherit",
                  padding: 0,
                }}
              />
            ) : (
              <span>{display}</span>
            )}
          </div>
        );
      }
    }
    return out;
  }, [sheet, styles, mergeIndex, startRow, endRow, startCol, endCol, selection, editing, onSelect, onCommitEdit]);

  // Bounding-box marquee — positioned over the union of the selection.
  let marquee: ReactNode = null;
  if (selection) {
    const n = normalizeSelection(selection);
    marquee = (
      <div
        data-testid="grid-marquee"
        aria-hidden
        style={{
          position: "absolute",
          top: HEADER_ROW_HEIGHT + n.r0 * ROW_HEIGHT,
          left: HEADER_COL_WIDTH + n.c0 * COL_WIDTH,
          width: (n.c1 - n.c0 + 1) * COL_WIDTH,
          height: (n.r1 - n.r0 + 1) * ROW_HEIGHT,
          border: "2px solid var(--ai-violet)",
          boxSizing: "border-box",
          pointerEvents: "none",
          zIndex: 4,
          background: isSingle(selection) ? "transparent" : "var(--ai-violet-light)",
          mixBlendMode: isSingle(selection) ? "normal" : "multiply",
        }}
      />
    );
  }

  // Column header band — tracks the viewport's top edge.
  const colHeaders: ReactNode[] = [];
  for (let c = startCol; c <= endCol; c++) {
    const n = selection ? normalizeSelection(selection) : null;
    const isActive = n ? c >= n.c0 && c <= n.c1 : false;
    colHeaders.push(
      <div
        key={`ch-${c}`}
        style={{
          position: "absolute",
          top: scroll.top,
          left: HEADER_COL_WIDTH + c * COL_WIDTH,
          width: COL_WIDTH,
          height: HEADER_ROW_HEIGHT,
          boxSizing: "border-box",
          borderRight: "1px solid var(--divider)",
          borderBottom: "1px solid var(--divider)",
          background: isActive ? "var(--hover)" : "var(--surface)",
          color: "var(--secondary)",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          fontSize: 11,
          fontWeight: 500,
          userSelect: "none",
          zIndex: 2,
        }}
      >
        {colToLetter(c)}
      </div>
    );
  }

  // Row header band — tracks the viewport's left edge.
  const rowHeaders: ReactNode[] = [];
  for (let r = startRow; r <= endRow; r++) {
    const n = selection ? normalizeSelection(selection) : null;
    const isActive = n ? r >= n.r0 && r <= n.r1 : false;
    rowHeaders.push(
      <div
        key={`rh-${r}`}
        style={{
          position: "absolute",
          top: HEADER_ROW_HEIGHT + r * ROW_HEIGHT,
          left: scroll.left,
          width: HEADER_COL_WIDTH,
          height: ROW_HEIGHT,
          boxSizing: "border-box",
          borderRight: "1px solid var(--divider)",
          borderBottom: "1px solid var(--divider)",
          background: isActive ? "var(--hover)" : "var(--surface)",
          color: "var(--secondary)",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          fontSize: 11,
          fontWeight: 500,
          userSelect: "none",
          zIndex: 2,
        }}
      >
        {r + 1}
      </div>
    );
  }

  return (
    <div
      ref={scrollRef}
      onScroll={onScroll}
      role="grid"
      aria-label={`Sheet ${sheet.name}`}
      data-testid="xlsx-grid"
      style={{
        position: "relative",
        width: "100%",
        height: "100%",
        overflow: "auto",
        background: "var(--background)",
        border: "1px solid var(--divider)",
        borderRadius: 6,
      }}
    >
      <div style={innerStyle}>
        {/* Top-left corner — anchored to the viewport edge. */}
        <div
          style={{
            position: "absolute",
            top: scroll.top,
            left: scroll.left,
            width: HEADER_COL_WIDTH,
            height: HEADER_ROW_HEIGHT,
            background: "var(--surface)",
            borderRight: "1px solid var(--divider)",
            borderBottom: "1px solid var(--divider)",
            zIndex: 3,
          }}
        />
        {colHeaders}
        {rowHeaders}
        {cellList}
        {marquee}
      </div>
    </div>
  );
}

// Re-export for convenience so XlsxEditor doesn't have to import from
// two adjacent modules.
export { singleSelection } from "./selection";
export type { CellPos, Selection } from "./selection";
