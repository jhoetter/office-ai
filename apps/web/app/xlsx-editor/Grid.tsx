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
 * Default cell geometry — used for any column / row that doesn't
 * carry an override in `sheet.columnWidths` / `sheet.rowHeights`.
 * Excel's defaults are 64px × 20px; we widen a touch so column
 * letters and short labels read better.
 */
const ROW_HEIGHT = 24;
const COL_WIDTH = 88;
const HEADER_ROW_HEIGHT = 24;
const HEADER_COL_WIDTH = 48;

const MIN_COL_WIDTH = 24;
const MIN_ROW_HEIGHT = 16;

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
  /**
   * Resize commit handlers. Called once on mouse-up after a header
   * drag. Parent dispatches `xlsx:set-column-width` /
   * `xlsx:set-row-height` so the same command-bus invariant holds
   * for resizing as for cell mutations.
   */
  readonly onResizeColumn?: (col: number, widthPx: number) => void;
  readonly onResizeRow?: (row: number, heightPx: number) => void;
  /**
   * Cells / ranges referenced by the formula currently being edited
   * in the formula bar (Phase 12c). Rendered as coloured borders so
   * the user can see what the formula points at — the colour matches
   * the same ref's coloured token in the formula bar.
   *
   * Only refs that resolve to the active sheet should be passed in;
   * the parent filters by sheet name before forwarding.
   */
  readonly refRects?: ReadonlyArray<RefRect>;
  /**
   * Selection commit handler used by the column / row header click
   * targets. The Grid raises which axis was clicked plus the index;
   * the parent maps it to a row-major or column-major selection.
   */
  readonly onSelectAxis?: (axis: "row" | "col", index: number, opts?: { extend?: boolean }) => void;
  /**
   * Right-click context-menu hooks. The Grid passes the click target
   * (variant + index/coords) plus the page coordinates the menu
   * should anchor to. Parent owns the menu state + entries so the
   * Grid stays presentational. (P13b)
   */
  readonly onContextMenu?: (target: GridContextTarget, coords: { x: number; y: number }) => void;
  /**
   * Marching-ants overlay for the active clipboard source range
   * (Phase 13d). When non-null, the Grid renders a dashed animated
   * border around the rectangle. `mode` distinguishes Cut from Copy
   * for the small visual hint that mirrors Excel.
   */
  readonly marchingAnts?: MarchingAntsRect | null;
  /**
   * Smart fill handle (Phase 13g). When the user drags the little
   * square at the bottom-right of the selection, the Grid manages
   * the whole interaction and raises this callback once on
   * mouse-up with the source rect, the extended target rect, and
   * the direction the user pulled in.
   */
  readonly onFill?: (args: {
    source: { r1: number; c1: number; r2: number; c2: number };
    target: { r1: number; c1: number; r2: number; c2: number };
    direction: "down" | "right" | "up" | "left";
  }) => void;
}

export interface MarchingAntsRect {
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
  readonly mode: "copy" | "cut";
}

export type GridContextTarget =
  | { readonly kind: "cell"; readonly row: number; readonly col: number }
  | { readonly kind: "row-header"; readonly row: number }
  | { readonly kind: "col-header"; readonly col: number };

export interface RefRect {
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
  readonly color: string;
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
  const {
    sheet,
    styles,
    selection,
    onSelect,
    onCommitEdit,
    onResizeColumn,
    onResizeRow,
    refRects,
    onSelectAxis,
    onContextMenu,
    marchingAnts,
    onFill,
  } = props;

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

  // Live drag state for column / row resize. While present, the
  // header bar shows the dragged size as a transient override that
  // is committed via `onResizeColumn` / `onResizeRow` on mouse-up.
  const [colDrag, setColDrag] = useState<{
    col: number;
    startX: number;
    startWidth: number;
    widthPx: number;
  } | null>(null);
  const [rowDrag, setRowDrag] = useState<{
    row: number;
    startY: number;
    startHeight: number;
    heightPx: number;
  } | null>(null);

  // Fill-handle drag state. `source` is captured at mouse-down from
  // the active selection; `preview` is recomputed from the cursor as
  // the user drags. The Grid manages this end-to-end and only
  // surfaces the final result through `onFill`.
  const [fillDrag, setFillDrag] = useState<{
    source: { r1: number; c1: number; r2: number; c2: number };
    preview: { r1: number; c1: number; r2: number; c2: number };
    direction: "down" | "right" | "up" | "left" | null;
  } | null>(null);

  // Global mousemove + mouseup for header resize. Tied to the
  // current `colDrag` / `rowDrag` so we drop the listeners when the
  // user isn't dragging.
  useEffect(() => {
    if (!colDrag) return;
    const onMove = (e: MouseEvent) => {
      const next = Math.max(MIN_COL_WIDTH, colDrag.startWidth + (e.clientX - colDrag.startX));
      setColDrag((prev) => (prev ? { ...prev, widthPx: next } : prev));
    };
    const onUp = () => {
      const finalPx = Math.max(MIN_COL_WIDTH, Math.round(colDrag.widthPx));
      setColDrag(null);
      onResizeColumn?.(colDrag.col, finalPx);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
  }, [colDrag, onResizeColumn]);

  useEffect(() => {
    if (!rowDrag) return;
    const onMove = (e: MouseEvent) => {
      const next = Math.max(MIN_ROW_HEIGHT, rowDrag.startHeight + (e.clientY - rowDrag.startY));
      setRowDrag((prev) => (prev ? { ...prev, heightPx: next } : prev));
    };
    const onUp = () => {
      const finalPx = Math.max(MIN_ROW_HEIGHT, Math.round(rowDrag.heightPx));
      setRowDrag(null);
      onResizeRow?.(rowDrag.row, finalPx);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
  }, [rowDrag, onResizeRow]);


  const onScroll = useCallback((ev: React.UIEvent<HTMLDivElement>) => {
    const el = ev.currentTarget;
    setScroll({ top: el.scrollTop, left: el.scrollLeft });
  }, []);

  // Compute the per-column width and per-row height with current
  // overrides + transient drag applied. Memoised on the inputs so
  // unchanged sheets reuse the prefix arrays (cheap, but it keeps
  // re-renders allocation-light).
  const { colWidths, rowHeights, colXs, rowYs, totalWidth, totalHeight } = useMemo(() => {
    const cw = new Array<number>(TOTAL_COLS);
    for (let c = 0; c < TOTAL_COLS; c++) {
      const override = sheet.columnWidths.get(c);
      let w = override ?? COL_WIDTH;
      if (colDrag && colDrag.col === c) w = colDrag.widthPx;
      cw[c] = Math.max(MIN_COL_WIDTH, w);
    }
    const rh = new Array<number>(TOTAL_ROWS);
    for (let r = 0; r < TOTAL_ROWS; r++) {
      const override = sheet.rowHeights.get(r);
      let h = override ?? ROW_HEIGHT;
      if (rowDrag && rowDrag.row === r) h = rowDrag.heightPx;
      rh[r] = Math.max(MIN_ROW_HEIGHT, h);
    }
    const cx = new Array<number>(TOTAL_COLS + 1);
    cx[0] = 0;
    for (let c = 0; c < TOTAL_COLS; c++) cx[c + 1] = cx[c]! + cw[c]!;
    const ry = new Array<number>(TOTAL_ROWS + 1);
    ry[0] = 0;
    for (let r = 0; r < TOTAL_ROWS; r++) ry[r + 1] = ry[r]! + rh[r]!;
    return {
      colWidths: cw,
      rowHeights: rh,
      colXs: cx,
      rowYs: ry,
      totalWidth: cx[TOTAL_COLS]!,
      totalHeight: ry[TOTAL_ROWS]!,
    };
  }, [sheet.columnWidths, sheet.rowHeights, colDrag, rowDrag]);

  const visibleH = Math.max(viewport.height - HEADER_ROW_HEIGHT, ROW_HEIGHT);
  const visibleW = Math.max(viewport.width - HEADER_COL_WIDTH, COL_WIDTH);

  // Binary-search the prefix arrays to find the first / last visible
  // index, then pad with the overscan window.
  const lower = (arr: ReadonlyArray<number>, target: number): number => {
    let lo = 0;
    let hi = arr.length - 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (arr[mid + 1]! <= target) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  };
  const startRow = Math.max(0, lower(rowYs, scroll.top) - OVERSCAN);
  const endRow = Math.min(TOTAL_ROWS - 1, lower(rowYs, scroll.top + visibleH) + OVERSCAN);
  const startCol = Math.max(0, lower(colXs, scroll.left) - OVERSCAN);
  const endCol = Math.min(TOTAL_COLS - 1, lower(colXs, scroll.left + visibleW) + OVERSCAN);

  // Global mouse handlers for the fill-handle drag. Lives here (after
  // colXs / rowYs are computed) so the cursor → cell hit-test sees
  // up-to-date geometry on every move event.
  useEffect(() => {
    if (!fillDrag) return;
    const el = scrollRef.current;
    if (!el) return;
    const onMove = (e: MouseEvent) => {
      const rect = el.getBoundingClientRect();
      const x = e.clientX - rect.left + el.scrollLeft - HEADER_COL_WIDTH;
      const y = e.clientY - rect.top + el.scrollTop - HEADER_ROW_HEIGHT;
      const cellRow = Math.max(0, Math.min(TOTAL_ROWS - 1, lower(rowYs, Math.max(0, y))));
      const cellCol = Math.max(0, Math.min(TOTAL_COLS - 1, lower(colXs, Math.max(0, x))));
      setFillDrag((prev) => {
        if (!prev) return prev;
        const src = prev.source;
        const dDown = Math.max(0, cellRow - src.r2);
        const dUp = Math.max(0, src.r1 - cellRow);
        const dRight = Math.max(0, cellCol - src.c2);
        const dLeft = Math.max(0, src.c1 - cellCol);
        const v = Math.max(dDown, dUp);
        const h = Math.max(dRight, dLeft);
        let direction: typeof prev.direction = null;
        let preview = src;
        if (v === 0 && h === 0) {
          preview = src;
        } else if (v >= h) {
          direction = dDown >= dUp ? "down" : "up";
          preview =
            direction === "down"
              ? { r1: src.r1, c1: src.c1, r2: cellRow, c2: src.c2 }
              : { r1: cellRow, c1: src.c1, r2: src.r2, c2: src.c2 };
        } else {
          direction = dRight >= dLeft ? "right" : "left";
          preview =
            direction === "right"
              ? { r1: src.r1, c1: src.c1, r2: src.r2, c2: cellCol }
              : { r1: src.r1, c1: cellCol, r2: src.r2, c2: src.c2 };
        }
        if (
          preview.r1 === prev.preview.r1 &&
          preview.c1 === prev.preview.c1 &&
          preview.r2 === prev.preview.r2 &&
          preview.c2 === prev.preview.c2 &&
          direction === prev.direction
        ) {
          return prev;
        }
        return { ...prev, preview, direction };
      });
    };
    const onUp = () => {
      setFillDrag((prev) => {
        if (prev && prev.direction && onFill) {
          onFill({ source: prev.source, target: prev.preview, direction: prev.direction });
        }
        return null;
      });
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
  }, [fillDrag, colXs, rowYs, onFill]);

  const innerStyle: CSSProperties = {
    position: "relative",
    width: HEADER_COL_WIDTH + totalWidth,
    height: HEADER_ROW_HEIGHT + totalHeight,
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
        const anchorCell = !!selection && selection.anchor.row === r && selection.anchor.col === c;
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
            onContextMenu={(e) => {
              if (!onContextMenu) return;
              e.preventDefault();
              // Excel parity: right-click on a cell *outside* the
              // current selection moves the selection there first;
              // right-click *inside* the selection keeps it. Either
              // way the menu sees the new selection because the
              // parent re-renders before the menu opens.
              if (!selection || !containsCell(selection, r, c)) {
                onSelect({ row: r, col: c });
              }
              onContextMenu({ kind: "cell", row: r, col: c }, { x: e.clientX, y: e.clientY });
            }}
            style={{
              position: "absolute",
              top: HEADER_ROW_HEIGHT + rowYs[r]!,
              left: HEADER_COL_WIDTH + colXs[c]!,
              width: colXs[c + widthCells]! - colXs[c]!,
              height: rowYs[r + heightCells]! - rowYs[r]!,
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
              ...(inSel && !anchorCell ? { background: "var(--ai-violet-light)" } : {}),
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
  }, [
    sheet,
    styles,
    mergeIndex,
    colXs,
    rowYs,
    startRow,
    endRow,
    startCol,
    endCol,
    selection,
    editing,
    onSelect,
    onCommitEdit,
    onContextMenu,
  ]);

  // Coloured borders for refs referenced by the formula currently
  // being edited (Phase 12c). Rendered behind the marquee so the
  // selection outline always wins on overlap.
  const refHighlights: ReactNode[] = [];
  if (refRects && refRects.length > 0) {
    for (let i = 0; i < refRects.length; i++) {
      const rect = refRects[i]!;
      const r0 = Math.max(0, Math.min(rect.r1, rect.r2));
      const r1 = Math.min(TOTAL_ROWS - 1, Math.max(rect.r1, rect.r2));
      const c0 = Math.max(0, Math.min(rect.c1, rect.c2));
      const c1 = Math.min(TOTAL_COLS - 1, Math.max(rect.c1, rect.c2));
      if (r0 > r1 || c0 > c1) continue;
      refHighlights.push(
        <div
          key={`refrect-${i}-${rect.color}-${r0}-${c0}-${r1}-${c1}`}
          data-testid={`ref-rect-${i}`}
          aria-hidden
          style={{
            position: "absolute",
            top: HEADER_ROW_HEIGHT + rowYs[r0]!,
            left: HEADER_COL_WIDTH + colXs[c0]!,
            width: colXs[c1 + 1]! - colXs[c0]!,
            height: rowYs[r1 + 1]! - rowYs[r0]!,
            border: `2px dashed ${rect.color}`,
            boxSizing: "border-box",
            pointerEvents: "none",
            zIndex: 3,
          }}
        />
      );
    }
  }

  // Marching-ants clipboard source overlay (Phase 13d). Drawn behind
  // the active selection marquee so the user can still see what's
  // currently selected; the dashed border keeps moving until the user
  // pastes or hits Escape (which the parent clears via the prop).
  let antsOverlay: ReactNode = null;
  if (marchingAnts) {
    const r0 = Math.max(0, Math.min(marchingAnts.r1, marchingAnts.r2));
    const r1 = Math.min(TOTAL_ROWS - 1, Math.max(marchingAnts.r1, marchingAnts.r2));
    const c0 = Math.max(0, Math.min(marchingAnts.c1, marchingAnts.c2));
    const c1 = Math.min(TOTAL_COLS - 1, Math.max(marchingAnts.c1, marchingAnts.c2));
    if (r0 <= r1 && c0 <= c1) {
      antsOverlay = (
        <div
          data-testid="grid-marching-ants"
          data-mode={marchingAnts.mode}
          aria-hidden
          className={`xlsx-marching-ants ${marchingAnts.mode === "cut" ? "xlsx-marching-ants-cut" : ""}`}
          style={{
            position: "absolute",
            top: HEADER_ROW_HEIGHT + rowYs[r0]!,
            left: HEADER_COL_WIDTH + colXs[c0]!,
            width: colXs[c1 + 1]! - colXs[c0]!,
            height: rowYs[r1 + 1]! - rowYs[r0]!,
            boxSizing: "border-box",
            pointerEvents: "none",
            zIndex: 5,
          }}
        />
      );
    }
  }

  // Bounding-box marquee — positioned over the union of the selection.
  let marquee: ReactNode = null;
  let fillHandle: ReactNode = null;
  if (selection) {
    const n = normalizeSelection(selection);
    marquee = (
      <div
        data-testid="grid-marquee"
        aria-hidden
        style={{
          position: "absolute",
          top: HEADER_ROW_HEIGHT + rowYs[n.r0]!,
          left: HEADER_COL_WIDTH + colXs[n.c0]!,
          width: colXs[n.c1 + 1]! - colXs[n.c0]!,
          height: rowYs[n.r1 + 1]! - rowYs[n.r0]!,
          border: "2px solid var(--ai-violet)",
          boxSizing: "border-box",
          pointerEvents: "none",
          zIndex: 4,
          background: isSingle(selection) ? "transparent" : "var(--ai-violet-light)",
          mixBlendMode: isSingle(selection) ? "normal" : "multiply",
        }}
      />
    );
    if (onFill) {
      const handleSize = 7;
      const right = HEADER_COL_WIDTH + colXs[n.c1 + 1]!;
      const bottom = HEADER_ROW_HEIGHT + rowYs[n.r1 + 1]!;
      fillHandle = (
        <div
          data-testid="grid-fill-handle"
          aria-label="Fill handle"
          role="button"
          onMouseDown={(e) => {
            e.preventDefault();
            e.stopPropagation();
            const src = { r1: n.r0, c1: n.c0, r2: n.r1, c2: n.c1 };
            setFillDrag({ source: src, preview: src, direction: null });
          }}
          style={{
            position: "absolute",
            top: bottom - Math.floor(handleSize / 2) - 1,
            left: right - Math.floor(handleSize / 2) - 1,
            width: handleSize,
            height: handleSize,
            background: "var(--ai-violet)",
            border: "1px solid white",
            cursor: "crosshair",
            zIndex: 6,
          }}
        />
      );
    }
  }

  let fillPreviewOverlay: ReactNode = null;
  if (fillDrag) {
    const fp = fillDrag.preview;
    const r0 = Math.max(0, Math.min(fp.r1, fp.r2));
    const r1 = Math.min(TOTAL_ROWS - 1, Math.max(fp.r1, fp.r2));
    const c0 = Math.max(0, Math.min(fp.c1, fp.c2));
    const c1 = Math.min(TOTAL_COLS - 1, Math.max(fp.c1, fp.c2));
    if (r0 <= r1 && c0 <= c1) {
      fillPreviewOverlay = (
        <div
          data-testid="grid-fill-preview"
          aria-hidden
          style={{
            position: "absolute",
            top: HEADER_ROW_HEIGHT + rowYs[r0]!,
            left: HEADER_COL_WIDTH + colXs[c0]!,
            width: colXs[c1 + 1]! - colXs[c0]!,
            height: rowYs[r1 + 1]! - rowYs[r0]!,
            border: "1px dashed var(--ai-violet)",
            boxSizing: "border-box",
            pointerEvents: "none",
            zIndex: 5,
          }}
        />
      );
    }
  }

  // Column header band — tracks the viewport's top edge.
  const colHeaders: ReactNode[] = [];
  for (let c = startCol; c <= endCol; c++) {
    const n = selection ? normalizeSelection(selection) : null;
    const isActive = n ? c >= n.c0 && c <= n.c1 : false;
    const colWidth = colWidths[c]!;
    colHeaders.push(
      <div
        key={`ch-${c}`}
        data-testid={`col-header-${colToLetter(c)}`}
        onMouseDown={(e) => {
          if (!onSelectAxis) return;
          // Ignore the resize-handle child — it stops propagation
          // already, but guard once more for clarity.
          const tgt = e.target as HTMLElement;
          if (tgt.dataset.testid?.startsWith("col-resize-")) return;
          e.preventDefault();
          onSelectAxis("col", c, { extend: e.shiftKey });
        }}
        onContextMenu={(e) => {
          if (!onContextMenu) return;
          e.preventDefault();
          // Right-click on a column header selects the whole
          // column first (Excel parity), then opens the menu.
          onSelectAxis?.("col", c);
          onContextMenu({ kind: "col-header", col: c }, { x: e.clientX, y: e.clientY });
        }}
        style={{
          position: "absolute",
          top: scroll.top,
          left: HEADER_COL_WIDTH + colXs[c]!,
          width: colWidth,
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
          cursor: onSelectAxis ? "pointer" : "default",
          zIndex: 2,
        }}
      >
        {colToLetter(c)}
        {onResizeColumn ? (
          <div
            data-testid={`col-resize-${colToLetter(c)}`}
            onMouseDown={(e) => {
              e.preventDefault();
              e.stopPropagation();
              setColDrag({
                col: c,
                startX: e.clientX,
                startWidth: colWidth,
                widthPx: colWidth,
              });
            }}
            style={{
              position: "absolute",
              top: 0,
              right: -3,
              width: 6,
              height: HEADER_ROW_HEIGHT,
              cursor: "col-resize",
              zIndex: 5,
            }}
          />
        ) : null}
      </div>
    );
  }

  // Row header band — tracks the viewport's left edge.
  const rowHeaders: ReactNode[] = [];
  for (let r = startRow; r <= endRow; r++) {
    const n = selection ? normalizeSelection(selection) : null;
    const isActive = n ? r >= n.r0 && r <= n.r1 : false;
    const rowHeight = rowHeights[r]!;
    rowHeaders.push(
      <div
        key={`rh-${r}`}
        data-testid={`row-header-${r + 1}`}
        onMouseDown={(e) => {
          if (!onSelectAxis) return;
          const tgt = e.target as HTMLElement;
          if (tgt.dataset.testid?.startsWith("row-resize-")) return;
          e.preventDefault();
          onSelectAxis("row", r, { extend: e.shiftKey });
        }}
        onContextMenu={(e) => {
          if (!onContextMenu) return;
          e.preventDefault();
          onSelectAxis?.("row", r);
          onContextMenu({ kind: "row-header", row: r }, { x: e.clientX, y: e.clientY });
        }}
        style={{
          position: "absolute",
          top: HEADER_ROW_HEIGHT + rowYs[r]!,
          left: scroll.left,
          width: HEADER_COL_WIDTH,
          height: rowHeight,
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
          cursor: onSelectAxis ? "pointer" : "default",
          zIndex: 2,
        }}
      >
        {r + 1}
        {onResizeRow ? (
          <div
            data-testid={`row-resize-${r + 1}`}
            onMouseDown={(e) => {
              e.preventDefault();
              e.stopPropagation();
              setRowDrag({
                row: r,
                startY: e.clientY,
                startHeight: rowHeight,
                heightPx: rowHeight,
              });
            }}
            style={{
              position: "absolute",
              left: 0,
              bottom: -3,
              height: 6,
              width: HEADER_COL_WIDTH,
              cursor: "row-resize",
              zIndex: 5,
            }}
          />
        ) : null}
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
        {refHighlights}
        {antsOverlay}
        {fillPreviewOverlay}
        {marquee}
        {fillHandle}
      </div>
    </div>
  );
}

// Re-export for convenience so XlsxEditor doesn't have to import from
// two adjacent modules.
export { singleSelection } from "./selection";
export type { CellPos, Selection } from "./selection";
