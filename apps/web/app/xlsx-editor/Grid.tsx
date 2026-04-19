"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { CSSProperties, ReactNode } from "react";
import { cellKey, colToLetter, type Sheet, type StyleTable } from "@officeai/xlsx";
import { ImageOverlay, type AnchorFromPx } from "./ImageOverlay";
import { ChartOverlay } from "./ChartOverlay";
import {
  containsCell,
  isSingle,
  normalizeSelection,
  singleSelection,
  type CellPos,
  type Selection,
} from "./selection";
import { formatCellValue, styleForCell } from "./styles";
import { buildGridDims, colXsView, rowYsView, type AxisLookup, type GridDims } from "./gridDimensions";

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
 * Total grid extents. C1 — Excel's true bounds. We never instantiate
 * dense per-cell arrays of this size: virtualisation paints only the
 * visible viewport, and {@link GridDims} resolves x/y positions in
 * O(log n) where n = number of column/row overrides (typically < 100).
 */
const TOTAL_ROWS = 1_048_576;
const TOTAL_COLS = 16_384;
const OVERSCAN = 4;

export type GridSelection = Selection;

export interface GridProps {
  readonly sheet: Sheet;
  /** Workbook style table — flattened per-cell to render fonts/fills/etc. */
  readonly styles: StyleTable;
  /**
   * C10 — sparse map of `r:c` → conditional-format overlay,
   * pre-computed by the parent so the Grid only has to apply it.
   * Empty / undefined when no rules are armed.
   */
  readonly cfOverlays?: ReadonlyMap<string, import("@officeai/xlsx").ConditionalFormatOverlay>;
  /**
   * C11 — sparse map of `r:c` → resolved list options for cells
   * covered by a typed `list` data-validation rule. The Grid uses
   * presence in this map to render the in-cell dropdown arrow and
   * pops a tiny picker when the arrow is clicked.
   */
  readonly dvIndex?: ReadonlyMap<string, ReadonlyArray<string>>;
  readonly selection: Selection | null;
  /**
   * C13 — Extra disjoint rectangles in the selection (Ctrl/Cmd-click).
   * The active area lives in `selection`; everything in here is just
   * additional marquee for Excel-parity multi-area selection.
   */
  readonly extraAreas?: ReadonlyArray<Selection>;
  /**
   * Drives selection moves. `extend` mirrors Shift-click / drag-extend:
   * keep `anchor`, replace `focus`. Single-cell mousedown leaves both
   * pinned to the same `pos`. `additive` (Ctrl/Cmd-click on Mac) starts
   * a new disjoint area without dropping the existing selection.
   */
  readonly onSelect: (pos: CellPos, opts?: { extend?: boolean; additive?: boolean }) => void;
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
  /**
   * Live mirror of the formula-bar draft for the active anchor cell
   * during type-to-edit / F2 / formula-bar editing. When non-null,
   * the matching cell renders the draft text (with a thin caret
   * indicator at the end) instead of the cell's committed display
   * value — Excel parity, where the cell visibly fills with the
   * keystrokes the user is typing even though the actual input
   * element lives in the formula bar.
   *
   * The double-click "in-cell editor" path is independent and takes
   * precedence: when the user is editing inside the cell directly,
   * its own `<input>` runs the show.
   */
  readonly liveEditDraft?: { readonly row: number; readonly col: number; readonly draft: string } | null;
  /**
   * Cells that carry an unresolved comment thread. Rendered with a
   * yellow border + soft fill behind the cell content (Word/Excel
   * convention) so the user can spot commented anchors at a glance.
   */
  readonly commentMarkers?: ReadonlyArray<CommentMarker>;
  /**
   * Imperative "scroll to + flash this cell" hook. When `nonce`
   * changes the Grid scrolls so the cell is visible and paints a
   * brief yellow pulse over it (~1.4 s) — used by the comments
   * sidebar's "click to locate" affordance. The parent owns the
   * nonce so repeated clicks on the same comment re-trigger the
   * effect.
   */
  readonly scrollTarget?: { readonly row: number; readonly col: number; readonly nonce: number } | null;
  /**
   * Click handler for the per-column AutoFilter dropdown chevron.
   * The Grid renders the chevron inside the header cell when
   * `sheet.autoFilter` covers that column; the parent owns the
   * dropdown's open/close state and dispatches the resulting
   * `xlsx:set-filter-column` / `xlsx:sort-range` commands.
   *
   * `colId` is the 0-based offset from `sheet.autoFilter.range.c1`.
   * `anchor` is the bounding rect of the chevron button (viewport
   * coords) so the dropdown can position itself.
   */
  readonly onOpenFilter?: (colId: number, anchor: DOMRect) => void;
  /**
   * Sheet-image overlays (Phase: image insertion). The Grid renders
   * each `sheet.images` entry as a free-floating overlay anchored
   * via `colXs / rowYs` so it tracks column / row resizes. Image
   * pointer interactions are routed through these callbacks; the
   * parent dispatches `xlsx:move-image` / `xlsx:resize-image`.
   */
  readonly imageObjectUrls?: ReadonlyMap<string, string>;
  readonly selectedImageId?: string | null;
  readonly onSelectImage?: (id: string | null) => void;
  readonly onMoveImage?: (id: string, anchor: AnchorFromPx) => void;
  readonly onResizeImage?: (id: string, size: { widthPx: number; heightPx: number }) => void;
  /**
   * Right-click on an image overlay surfaces a dedicated context
   * target so the parent can show "Delete image" entries. When the
   * handler isn't wired, image right-clicks fall through to the
   * cell context menu underneath (Excel parity).
   */
  readonly onImageContextMenu?: (imageId: string, coords: { x: number; y: number }) => void;
  /**
   * C15 — chart overlays. The Grid pins each `sheet.charts` entry
   * to its anchor (same coordinate model as images). Selection +
   * removal route through the parent's `xlsx:remove-chart` /
   * `xlsx:move-chart` handlers; v1 has no inline drag.
   */
  readonly selectedChartId?: string | null;
  readonly onSelectChart?: (id: string | null) => void;
  readonly onRemoveChart?: (id: string) => void;
}

export interface CommentMarker {
  readonly row: number;
  readonly col: number;
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
  | { readonly kind: "col-header"; readonly col: number }
  | { readonly kind: "image"; readonly imageId: string };

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
    liveEditDraft,
    commentMarkers,
    scrollTarget,
    onOpenFilter,
    imageObjectUrls,
    selectedImageId,
    onSelectImage,
    onMoveImage,
    onResizeImage,
    onImageContextMenu,
    cfOverlays,
    dvIndex,
    extraAreas,
    selectedChartId,
    onSelectChart,
    onRemoveChart,
  } = props;
  const extras: ReadonlyArray<Selection> = extraAreas ?? [];

  const scrollRef = useRef<HTMLDivElement | null>(null);
  const [scroll, setScroll] = useState({ top: 0, left: 0 });
  // Transient yellow pulse painted over a cell after the comments
  // sidebar issues a "scroll to me" request. Cleared by a timer so
  // the highlight fades on its own.
  const [flashCell, setFlashCell] = useState<{ row: number; col: number } | null>(null);
  const [editing, setEditing] = useState<{ row: number; col: number; draft: string } | null>(null);
  // C11 — Open data-validation list popover for the named cell.
  const [dvPicker, setDvPicker] = useState<{ row: number; col: number } | null>(null);
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
  type FillDragState = {
    source: { r1: number; c1: number; r2: number; c2: number };
    preview: { r1: number; c1: number; r2: number; c2: number };
    direction: "down" | "right" | "up" | "left" | null;
  };
  const [fillDrag, setFillDrag] = useState<FillDragState | null>(null);
  // Mirror of the latest fillDrag so the global mouseup handler can
  // read the freshest preview/direction without having to live inside
  // a setState updater (where dispatching onFill — which ultimately
  // fires setSnapshot on the parent — would trigger React's
  // "setState while rendering" warning).
  const fillDragRef = useRef<FillDragState | null>(null);
  useEffect(() => {
    fillDragRef.current = fillDrag;
  }, [fillDrag]);

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

  // C1 — Real Excel bounds (1,048,576 × 16,384) demand an
  // index-backed dimension model: dense prefix sums over a million
  // rows would burn ~10 ms per re-render. Instead we fold the sparse
  // override maps + transient drag state into a single `GridDims`
  // object that resolves any `xAt(c)` / `yAt(r)` query in O(log n)
  // where n = number of column/row overrides (typically < 100).
  const dims: GridDims = useMemo(() => {
    const colOverrides = new Map<number, number>(sheet.columnWidths);
    if (colDrag) {
      colOverrides.set(colDrag.col, Math.max(MIN_COL_WIDTH, colDrag.widthPx));
    }
    const rowOverrides = new Map<number, number>(sheet.rowHeights);
    for (const r of sheet.hiddenRows) rowOverrides.set(r, 0);
    if (rowDrag) {
      rowOverrides.set(rowDrag.row, Math.max(MIN_ROW_HEIGHT, rowDrag.heightPx));
    }
    return buildGridDims({
      columnWidths: colOverrides,
      rowHeights: rowOverrides,
      defaultColWidth: COL_WIDTH,
      defaultRowHeight: ROW_HEIGHT,
      totalRows: TOTAL_ROWS,
      totalCols: TOTAL_COLS,
    });
  }, [sheet.columnWidths, sheet.rowHeights, sheet.hiddenRows, colDrag, rowDrag]);

  // Array-like proxies preserve the existing `colXs[c]` access shape.
  // Reads delegate to `dims.xAt(c)` / `dims.yAt(r)` so the rest of
  // the file stays readable without a 50-callsite rewrite.
  const colXs: AxisLookup = useMemo(() => colXsView(dims), [dims]);
  const rowYs: AxisLookup = useMemo(() => rowYsView(dims), [dims]);
  const totalWidth = dims.totalWidth;
  const totalHeight = dims.totalHeight;

  const visibleH = Math.max(viewport.height - HEADER_ROW_HEIGHT, ROW_HEIGHT);
  const visibleW = Math.max(viewport.width - HEADER_COL_WIDTH, COL_WIDTH);

  // Whenever the parent bumps `scrollTarget.nonce` we scroll the
  // requested cell into view (with a small margin so the cell isn't
  // pinned against the headers) and trigger a yellow flash overlay
  // that auto-clears.
  const scrollTargetNonce = scrollTarget?.nonce ?? null;
  const scrollTargetRow = scrollTarget?.row;
  const scrollTargetCol = scrollTarget?.col;
  useEffect(() => {
    if (scrollTargetNonce == null) return;
    if (scrollTargetRow == null || scrollTargetCol == null) return;
    const el = scrollRef.current;
    if (!el) return;
    const cellLeft = colXs[scrollTargetCol] ?? 0;
    const cellRight = colXs[scrollTargetCol + 1] ?? cellLeft;
    const cellTop = rowYs[scrollTargetRow] ?? 0;
    const cellBottom = rowYs[scrollTargetRow + 1] ?? cellTop;
    const viewLeft = el.scrollLeft;
    const viewRight = viewLeft + Math.max(el.clientWidth - HEADER_COL_WIDTH, 0);
    const viewTop = el.scrollTop;
    const viewBottom = viewTop + Math.max(el.clientHeight - HEADER_ROW_HEIGHT, 0);
    let nextLeft = viewLeft;
    let nextTop = viewTop;
    const margin = 16;
    if (cellLeft < viewLeft) nextLeft = Math.max(0, cellLeft - margin);
    else if (cellRight > viewRight) nextLeft = Math.max(0, cellRight - (viewRight - viewLeft) + margin);
    if (cellTop < viewTop) nextTop = Math.max(0, cellTop - margin);
    else if (cellBottom > viewBottom) nextTop = Math.max(0, cellBottom - (viewBottom - viewTop) + margin);
    if (nextLeft !== viewLeft || nextTop !== viewTop) {
      el.scrollTo({ left: nextLeft, top: nextTop, behavior: "smooth" });
    }
    setFlashCell({ row: scrollTargetRow, col: scrollTargetCol });
    const handle = window.setTimeout(() => setFlashCell(null), 1600);
    return () => window.clearTimeout(handle);
  }, [scrollTargetNonce, scrollTargetRow, scrollTargetCol, colXs, rowYs]);

  // C1 — derive the visible window via the lazy index. `colAtX` /
  // `rowAtY` resolve to the index whose left/top edge is the largest
  // ≤ the given pixel position, exactly what the old prefix-array
  // binary search returned.
  const startRow = Math.max(0, dims.rowAtY(scroll.top) - OVERSCAN);
  const endRow = Math.min(TOTAL_ROWS - 1, dims.rowAtY(scroll.top + visibleH) + OVERSCAN);
  const startCol = Math.max(0, dims.colAtX(scroll.left) - OVERSCAN);
  const endCol = Math.min(TOTAL_COLS - 1, dims.colAtX(scroll.left + visibleW) + OVERSCAN);

  // C3 — Freeze panes. The first `freezeRows` rows / `freezeCols`
  // columns stay pinned to the viewport edge as the user scrolls,
  // mirroring Excel's classic `View > Freeze Panes`. The
  // implementation uses scroll-offset positioning rather than four
  // separate scroll containers: a pinned cell's DOM `top` is set to
  // `scroll.top + …` so that the inner container's scroll-induced
  // upward shift is exactly cancelled out, leaving the cell at a
  // fixed viewport position. The trade-off is that the "scrolled"
  // cells render *behind* the pinned ones (z-index 0 vs 2) and rely
  // on the frozen cells' opaque background to hide them. That's
  // visually equivalent for a 2D viewport and avoids the
  // synchronised-scroll dance a true 4-quadrant DOM would need.
  const freezeRows = Math.min(sheet.freeze?.rows ?? 0, TOTAL_ROWS);
  const freezeCols = Math.min(sheet.freeze?.cols ?? 0, TOTAL_COLS);
  const hasFreeze = freezeRows > 0 || freezeCols > 0;
  const freezeYPx = freezeRows > 0 ? dims.yAt(freezeRows) : 0;
  const freezeXPx = freezeCols > 0 ? dims.xAt(freezeCols) : 0;

  const topFor = useCallback(
    (r: number): number => HEADER_ROW_HEIGHT + (rowYs[r] ?? 0) + (r < freezeRows ? scroll.top : 0),
    [rowYs, freezeRows, scroll.top]
  );
  const leftFor = useCallback(
    (c: number): number => HEADER_COL_WIDTH + (colXs[c] ?? 0) + (c < freezeCols ? scroll.left : 0),
    [colXs, freezeCols, scroll.left]
  );

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
      const cellRow = Math.max(0, Math.min(TOTAL_ROWS - 1, dims.rowAtY(Math.max(0, y))));
      const cellCol = Math.max(0, Math.min(TOTAL_COLS - 1, dims.colAtX(Math.max(0, x))));
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
      // Read the latest drag state from the ref (kept in sync via the
      // effect above), then clear it via setState. `onFill` runs
      // *outside* any setState updater so the downstream
      // `applyCommand` -> `subscribe` -> `setSnapshot` chain on the
      // parent doesn't fire during Grid reconciliation, which React
      // flags as "setState while rendering a different component".
      const finalDrag = fillDragRef.current;
      setFillDrag(null);
      if (finalDrag && finalDrag.direction && onFill) {
        onFill({
          source: finalDrag.source,
          target: finalDrag.preview,
          direction: finalDrag.direction,
        });
      }
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
  }, [fillDrag, dims, onFill]);

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

  // Build the union of frozen + visible rows / cols so we render each
  // cell exactly once. Frozen rows/cols are always present; the
  // unfrozen visible window starts past the freeze split so we don't
  // double-paint the same coordinate in two quadrants.
  const rowsToRender: number[] = useMemo(() => {
    const out: number[] = [];
    for (let r = 0; r < freezeRows; r++) out.push(r);
    const start = Math.max(freezeRows, startRow);
    for (let r = start; r <= endRow; r++) out.push(r);
    return out;
  }, [freezeRows, startRow, endRow]);

  const colsToRender: number[] = useMemo(() => {
    const out: number[] = [];
    for (let c = 0; c < freezeCols; c++) out.push(c);
    const start = Math.max(freezeCols, startCol);
    for (let c = start; c <= endCol; c++) out.push(c);
    return out;
  }, [freezeCols, startCol, endCol]);

  const cellList: ReactNode[] = useMemo(() => {
    const out: ReactNode[] = [];
    for (const r of rowsToRender) {
      if (sheet.hiddenRows.has(r)) continue;
      const rFrozen = r < freezeRows;
      for (const c of colsToRender) {
        const cFrozen = c < freezeCols;
        const cellFrozen = rFrozen || cFrozen;
        const k = cellKey(r, c);
        if (mergeIndex.covered.has(k)) continue;
        const merge = mergeIndex.topLeft.get(k);
        const widthCells = merge ? merge.c2 - merge.c1 + 1 : 1;
        const heightCells = merge ? merge.r2 - merge.r1 + 1 : 1;
        const cell = sheet.cells.get(k);
        const cellStyle = styleForCell(styles, cell?.styleId);
        const display = cell ? formatCellValue(cell.value, cellStyle.effective.numFmtId) : "";
        const cfOverlay = cfOverlays?.get(k);
        const inSel =
          (!!selection && containsCell(selection, r, c)) || extras.some((a) => containsCell(a, r, c));
        const anchorCell = !!selection && selection.anchor.row === r && selection.anchor.col === c;
        const isEditing = editing?.row === r && editing?.col === c;
        // Live draft from the formula bar mirrored into the cell
        // during type-to-edit. Only kicks in when the in-cell editor
        // (`isEditing`) isn't already running the show — the in-cell
        // editor owns the cell's contents while it's open.
        const isLiveDrafting =
          !isEditing && !!liveEditDraft && liveEditDraft.row === r && liveEditDraft.col === c;
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
              const additive = (e.metaKey || e.ctrlKey) && !e.shiftKey;
              onSelect({ row: r, col: c }, { extend: e.shiftKey, additive });
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
              top: topFor(r),
              left: leftFor(c),
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
              // Frozen cells sit above scrolled cells so their opaque
              // background hides the scrolled rows/cols underneath. The
              // anchor cell stays one tier above its row's z-index so
              // the white anchor pip beats the violet selection wash.
              zIndex: cellFrozen ? (anchorCell ? 3 : 2) : anchorCell ? 1 : 0,
              ...cellStyle.css,
              // C10 — conditional-format overlay. Layers on top of
              // the base cellStyle.css so a CF rule that paints
              // `fill` or `fontColor` wins over the cell's stored
              // style, but unrelated style fields (numFmt, borders,
              // alignment, font family/size) survive. Selection
              // background still wins below for non-anchor cells.
              ...(cfOverlay?.fill ? { background: `#${cfOverlay.fill}` } : {}),
              ...(cfOverlay?.fontColor ? { color: `#${cfOverlay.fontColor}` } : {}),
              ...(cfOverlay?.bold ? { fontWeight: 700 } : {}),
              ...(cfOverlay?.italic ? { fontStyle: "italic" as const } : {}),
              // Selection background wins over a per-cell fill so the
              // user can still see what's highlighted, except for the
              // anchor where Excel keeps the cell's real fill.
              ...(inSel && !anchorCell ? { background: "var(--ai-violet-light)" } : {}),
            }}
          >
            {cfOverlay?.barColor && cfOverlay.barFraction !== undefined && !isEditing ? (
              <div
                aria-hidden
                style={{
                  position: "absolute",
                  inset: 0,
                  background: `linear-gradient(to right, #${cfOverlay.barColor}33 0%, #${cfOverlay.barColor}33 ${cfOverlay.barFraction * 100}%, transparent ${cfOverlay.barFraction * 100}%)`,
                  pointerEvents: "none",
                }}
              />
            ) : null}
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
            ) : isLiveDrafting ? (
              <span
                data-testid={`cell-${colToLetter(c)}${r + 1}-live-draft`}
                style={{ display: "inline-flex", alignItems: "center" }}
              >
                <span style={{ color: "var(--foreground)" }}>{liveEditDraft.draft}</span>
                <span
                  aria-hidden
                  className="xlsx-cell-caret"
                  style={{
                    display: "inline-block",
                    width: 1,
                    height: "1em",
                    marginLeft: 1,
                    background: "var(--foreground)",
                  }}
                />
              </span>
            ) : (
              <span
                style={{ display: "inline-block", flex: 1, overflow: "hidden", textOverflow: "ellipsis" }}
              >
                {display}
              </span>
            )}
            {/* C11 — In-cell list dropdown arrow. Only painted when
                the cell carries a typed `list` data-validation rule
                AND the cell is currently the anchor (Excel parity:
                only the active cell in a validated range shows the
                arrow). */}
            {dvIndex && dvIndex.has(k) && anchorCell && !isEditing ? (
              <button
                type="button"
                data-testid={`dv-open-${colToLetter(c)}${r + 1}`}
                aria-label="Open data-validation list"
                title="Open list"
                onMouseDown={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                }}
                onClick={(e) => {
                  e.stopPropagation();
                  setDvPicker({ row: r, col: c });
                }}
                style={{
                  position: "absolute",
                  right: -1,
                  top: 0,
                  bottom: 0,
                  width: 18,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  background: "var(--surface)",
                  border: "1px solid var(--divider)",
                  borderRadius: 2,
                  cursor: "pointer",
                  fontSize: 9,
                  color: "var(--foreground)",
                  zIndex: 4,
                }}
              >
                ▾
              </button>
            ) : null}
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
    rowsToRender,
    colsToRender,
    freezeRows,
    freezeCols,
    topFor,
    leftFor,
    selection,
    editing,
    onSelect,
    onCommitEdit,
    onContextMenu,
    liveEditDraft,
    cfOverlays,
    dvIndex,
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

  // Unresolved-comment markers. Rendered behind the selection marquee
  // so the marquee always wins on overlap, but in front of cell content
  // so the soft yellow fill reads through to the user. We also render a
  // tiny triangular tag at the top-right of the cell — Excel parity.
  const commentOverlays: ReactNode[] = [];
  if (commentMarkers && commentMarkers.length > 0) {
    for (let i = 0; i < commentMarkers.length; i++) {
      const m = commentMarkers[i]!;
      if (m.row < 0 || m.row >= TOTAL_ROWS) continue;
      if (m.col < 0 || m.col >= TOTAL_COLS) continue;
      const top = HEADER_ROW_HEIGHT + rowYs[m.row]!;
      const left = HEADER_COL_WIDTH + colXs[m.col]!;
      const width = colXs[m.col + 1]! - colXs[m.col]!;
      const height = rowYs[m.row + 1]! - rowYs[m.row]!;
      commentOverlays.push(
        <div
          key={`comment-cell-${m.row}-${m.col}`}
          data-testid={`comment-marker-${m.row}-${m.col}`}
          aria-hidden
          style={{
            position: "absolute",
            top,
            left,
            width,
            height,
            border: "1.5px solid var(--warning)",
            background: "color-mix(in srgb, var(--warning) 18%, transparent)",
            boxSizing: "border-box",
            pointerEvents: "none",
            zIndex: 3,
          }}
        />
      );
      commentOverlays.push(
        <div
          key={`comment-tag-${m.row}-${m.col}`}
          aria-hidden
          style={{
            position: "absolute",
            top,
            left: left + width - 7,
            width: 0,
            height: 0,
            borderTop: "7px solid var(--warning)",
            borderLeft: "7px solid transparent",
            pointerEvents: "none",
            zIndex: 3,
          }}
        />
      );
    }
  }

  // Brief yellow pulse painted on top of a cell after a comments-rail
  // "scroll to me" request. Uses a CSS animation defined in
  // `globals.css` (`.xlsx-comment-flash`) so the highlight fades on
  // its own without re-renders.
  let flashOverlay: ReactNode = null;
  if (
    flashCell &&
    flashCell.row >= 0 &&
    flashCell.row < TOTAL_ROWS &&
    flashCell.col >= 0 &&
    flashCell.col < TOTAL_COLS
  ) {
    const ftop = HEADER_ROW_HEIGHT + rowYs[flashCell.row]!;
    const fleft = HEADER_COL_WIDTH + colXs[flashCell.col]!;
    const fwidth = colXs[flashCell.col + 1]! - colXs[flashCell.col]!;
    const fheight = rowYs[flashCell.row + 1]! - rowYs[flashCell.row]!;
    flashOverlay = (
      <div
        key={`flash-${flashCell.row}-${flashCell.col}`}
        data-testid="comment-flash"
        aria-hidden
        className="xlsx-comment-flash"
        style={{
          position: "absolute",
          top: ftop,
          left: fleft,
          width: fwidth,
          height: fheight,
          boxSizing: "border-box",
          pointerEvents: "none",
          zIndex: 5,
        }}
      />
    );
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

  // C14 — Excel Tables: outer border + header band + zebra row banding
  // overlay. Sits beneath the selection marquee but above cells so the
  // table style is visible without disturbing per-cell formatting.
  const tableOverlays: ReactNode[] = [];
  if (sheet.tables.length > 0) {
    for (const table of sheet.tables) {
      const parsed = parseTableA1(table.range);
      if (!parsed) continue;
      const tr0 = parsed.r1;
      const tr1 = parsed.r2;
      const tc0 = parsed.c1;
      const tc1 = parsed.c2;
      if (tr1 >= rowYs.length - 1 || tc1 >= colXs.length - 1) continue;
      const top = HEADER_ROW_HEIGHT + rowYs[tr0]!;
      const left = HEADER_COL_WIDTH + colXs[tc0]!;
      const width = colXs[tc1 + 1]! - colXs[tc0]!;
      const height = rowYs[tr1 + 1]! - rowYs[tr0]!;
      tableOverlays.push(
        <div
          key={`table-outline-${table.tableId}`}
          data-testid={`grid-table-outline-${table.tableId}`}
          aria-hidden
          style={{
            position: "absolute",
            top,
            left,
            width,
            height,
            border: "1px solid var(--ai-violet)",
            boxSizing: "border-box",
            pointerEvents: "none",
            zIndex: 1,
          }}
        />
      );
      if (table.headerRowCount > 0) {
        const headerEndRow = Math.min(tr1, tr0 + table.headerRowCount - 1);
        const headerHeight = rowYs[headerEndRow + 1]! - rowYs[tr0]!;
        tableOverlays.push(
          <div
            key={`table-header-${table.tableId}`}
            data-testid={`grid-table-header-${table.tableId}`}
            aria-hidden
            style={{
              position: "absolute",
              top,
              left,
              width,
              height: headerHeight,
              background: "var(--ai-violet-light)",
              borderBottom: "1px solid var(--ai-violet)",
              boxSizing: "border-box",
              pointerEvents: "none",
              zIndex: 1,
              opacity: 0.6,
              mixBlendMode: "multiply",
            }}
          />
        );
      }
      // Zebra body banding — paint every other body row with a faint
      // tint. Excel's default Table style does the same.
      const bodyStart = tr0 + table.headerRowCount;
      for (let r = bodyStart; r <= tr1; r++) {
        if ((r - bodyStart) % 2 === 0) continue;
        if (sheet.hiddenRows.has(r)) continue;
        const rowTop = HEADER_ROW_HEIGHT + rowYs[r]!;
        const rowHeight = rowYs[r + 1]! - rowYs[r]!;
        tableOverlays.push(
          <div
            key={`table-band-${table.tableId}-${r}`}
            aria-hidden
            style={{
              position: "absolute",
              top: rowTop,
              left,
              width,
              height: rowHeight,
              background: "var(--muted)",
              opacity: 0.35,
              pointerEvents: "none",
              zIndex: 1,
              mixBlendMode: "multiply",
            }}
          />
        );
      }
    }
  }

  // C13 — Extra disjoint areas (Ctrl-click). Same look as the active
  // marquee minus the fill handle; a slightly thinner border keeps the
  // active area visually dominant so the user knows where the next
  // commit / paste / fill will land.
  const extraAreaOverlays: ReactNode[] = [];
  for (let i = 0; i < extras.length; i++) {
    const a = extras[i]!;
    const n = normalizeSelection(a);
    extraAreaOverlays.push(
      <div
        key={`extra-area-${i}`}
        data-testid={`grid-extra-area-${i}`}
        aria-hidden
        style={{
          position: "absolute",
          top: HEADER_ROW_HEIGHT + rowYs[n.r0]!,
          left: HEADER_COL_WIDTH + colXs[n.c0]!,
          width: colXs[n.c1 + 1]! - colXs[n.c0]!,
          height: rowYs[n.r1 + 1]! - rowYs[n.r0]!,
          border: "1.5px solid var(--ai-violet)",
          boxSizing: "border-box",
          pointerEvents: "none",
          zIndex: 4,
          background: "var(--ai-violet-light)",
          mixBlendMode: "multiply",
          opacity: 0.85,
        }}
      />
    );
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

  // Column header band — tracks the viewport's top edge. Frozen
  // columns also pin horizontally so their header letter stays in
  // sync with the pinned cells underneath.
  const colHeaders: ReactNode[] = [];
  const af = sheet.autoFilter;
  for (const c of colsToRender) {
    const n = selection ? normalizeSelection(selection) : null;
    const isActive = n ? c >= n.c0 && c <= n.c1 : false;
    const colWidth = dims.colWidth(c);
    const cFrozen = c < freezeCols;
    const filterColId = af && c >= af.range.c1 && c <= af.range.c2 ? c - af.range.c1 : null;
    const filterActive = filterColId !== null && af!.columns.has(filterColId);
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
          left: HEADER_COL_WIDTH + colXs[c]! + (cFrozen ? scroll.left : 0),
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
          zIndex: cFrozen ? 4 : 2,
        }}
      >
        {colToLetter(c)}
        {filterColId !== null && onOpenFilter ? (
          <button
            type="button"
            data-testid={`col-filter-${colToLetter(c)}`}
            aria-label={`Filter column ${colToLetter(c)}`}
            onMouseDown={(e) => {
              e.preventDefault();
              e.stopPropagation();
            }}
            onClick={(e) => {
              e.preventDefault();
              e.stopPropagation();
              const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
              onOpenFilter(filterColId, rect);
            }}
            style={{
              position: "absolute",
              right: 4,
              top: "50%",
              transform: "translateY(-50%)",
              width: 14,
              height: 14,
              padding: 0,
              border: "1px solid var(--divider)",
              borderRadius: 2,
              background: filterActive ? "var(--accent, #2563eb)" : "var(--surface)",
              color: filterActive ? "#fff" : "var(--secondary)",
              fontSize: 9,
              lineHeight: 1,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              cursor: "pointer",
              zIndex: 6,
            }}
          >
            ▾
          </button>
        ) : null}
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

  // Row header band — tracks the viewport's left edge. Frozen rows
  // also pin vertically so their numeric label stays aligned with the
  // pinned cells alongside.
  const rowHeaders: ReactNode[] = [];
  for (const r of rowsToRender) {
    if (sheet.hiddenRows.has(r)) continue;
    const n = selection ? normalizeSelection(selection) : null;
    const isActive = n ? r >= n.r0 && r <= n.r1 : false;
    const rowHeight = dims.rowHeight(r);
    const rFrozen = r < freezeRows;
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
          top: HEADER_ROW_HEIGHT + rowYs[r]! + (rFrozen ? scroll.top : 0),
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
          zIndex: rFrozen ? 4 : 2,
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
        {sheet.charts.map((chart) => {
          const a = chart.anchor;
          const top = HEADER_ROW_HEIGHT + (rowYs[a.fromRow] ?? 0) + a.fromOffsetYPx;
          const left = HEADER_COL_WIDTH + (colXs[a.fromCol] ?? 0) + a.fromOffsetXPx;
          return (
            <div
              key={`chart-wrap-${chart.id}`}
              style={{ position: "absolute", top, left, zIndex: 5 }}
              onMouseDown={(e) => {
                e.stopPropagation();
                onSelectChart?.(chart.id);
              }}
            >
              <ChartOverlay
                chart={chart}
                sheet={sheet}
                width={a.widthPx}
                height={a.heightPx}
                selected={selectedChartId === chart.id}
                onSelect={() => onSelectChart?.(chart.id)}
                onRequestRemove={onRemoveChart ? () => onRemoveChart(chart.id) : undefined}
              />
            </div>
          );
        })}
        {sheet.images.map((image) => (
          <div
            key={`image-wrap-${image.id}`}
            onContextMenu={(e) => {
              if (!onImageContextMenu) return;
              e.preventDefault();
              e.stopPropagation();
              onSelectImage?.(image.id);
              onImageContextMenu(image.id, { x: e.clientX, y: e.clientY });
            }}
          >
            <ImageOverlay
              image={image}
              imageId={image.id}
              src={imageObjectUrls?.get(image.mediaRef)}
              colXs={colXs}
              rowYs={rowYs}
              headerOffset={{ x: HEADER_COL_WIDTH, y: HEADER_ROW_HEIGHT }}
              selected={selectedImageId === image.id}
              onSelect={() => onSelectImage?.(image.id)}
              onMoveCommit={(anchor) => onMoveImage?.(image.id, anchor)}
              onResizeCommit={(size) => onResizeImage?.(image.id, size)}
            />
          </div>
        ))}
        {refHighlights}
        {commentOverlays}
        {flashOverlay}
        {antsOverlay}
        {fillPreviewOverlay}
        {tableOverlays}
        {extraAreaOverlays}
        {marquee}
        {fillHandle}
        {hasFreeze && freezeRows > 0 ? (
          <div
            data-testid="freeze-divider-h"
            aria-hidden
            style={{
              position: "absolute",
              top: scroll.top + HEADER_ROW_HEIGHT + freezeYPx - 1,
              left: scroll.left,
              width: viewport.width,
              height: 2,
              background: "var(--ai-violet)",
              opacity: 0.6,
              pointerEvents: "none",
              zIndex: 5,
            }}
          />
        ) : null}
        {hasFreeze && freezeCols > 0 ? (
          <div
            data-testid="freeze-divider-v"
            aria-hidden
            style={{
              position: "absolute",
              top: scroll.top,
              left: scroll.left + HEADER_COL_WIDTH + freezeXPx - 1,
              width: 2,
              height: viewport.height,
              background: "var(--ai-violet)",
              opacity: 0.6,
              pointerEvents: "none",
              zIndex: 5,
            }}
          />
        ) : null}
        {dvPicker && dvIndex
          ? (() => {
              const opts = dvIndex.get(cellKey(dvPicker.row, dvPicker.col));
              if (!opts) return null;
              return (
                <DvListPicker
                  options={opts}
                  top={topFor(dvPicker.row) + (rowYs[dvPicker.row + 1]! - rowYs[dvPicker.row]!)}
                  left={leftFor(dvPicker.col)}
                  width={Math.max(120, colXs[dvPicker.col + 1]! - colXs[dvPicker.col]!)}
                  onPick={(value) => {
                    onCommitEdit({ row: dvPicker.row, col: dvPicker.col }, value);
                    setDvPicker(null);
                  }}
                  onClose={() => setDvPicker(null)}
                />
              );
            })()
          : null}
      </div>
    </div>
  );
}

/**
 * Tiny A1-range parser for the Tables overlay. We deliberately avoid
 * depending on the heavier `parseRange` from `@officeai/xlsx` here so
 * the Grid stays free of cross-package side-effects on every render;
 * the regex covers the only shape we ever store on `TableDef.range`
 * (`A1:Z99`-style, 0-based output).
 */
function parseTableA1(range: string): { r1: number; c1: number; r2: number; c2: number } | null {
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(range);
  if (!m) return null;
  const c1 = letterToColIndex(m[1]!);
  const r1 = Number.parseInt(m[2]!, 10) - 1;
  const c2 = letterToColIndex(m[3]!);
  const r2 = Number.parseInt(m[4]!, 10) - 1;
  if (![r1, c1, r2, c2].every(Number.isFinite)) return null;
  return {
    r1: Math.min(r1, r2),
    c1: Math.min(c1, c2),
    r2: Math.max(r1, r2),
    c2: Math.max(c1, c2),
  };
}

function letterToColIndex(letter: string): number {
  let n = 0;
  for (const ch of letter) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

interface DvListPickerProps {
  readonly options: ReadonlyArray<string>;
  readonly top: number;
  readonly left: number;
  readonly width: number;
  readonly onPick: (value: string) => void;
  readonly onClose: () => void;
}

function DvListPicker(props: DvListPickerProps): ReactNode {
  const { options, top, left, width, onPick, onClose } = props;
  useEffect(() => {
    const onDoc = (e: MouseEvent) => {
      const t = e.target as HTMLElement | null;
      if (t && t.closest("[data-dv-picker]")) return;
      onClose();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        onClose();
      }
    };
    window.addEventListener("mousedown", onDoc);
    window.addEventListener("keydown", onKey);
    return () => {
      window.removeEventListener("mousedown", onDoc);
      window.removeEventListener("keydown", onKey);
    };
  }, [onClose]);

  return (
    <div
      data-dv-picker
      data-testid="dv-list-picker"
      role="listbox"
      aria-label="Data validation list"
      style={{
        position: "absolute",
        top,
        left,
        width,
        maxHeight: 220,
        overflowY: "auto",
        background: "var(--background)",
        border: "1px solid var(--divider)",
        borderRadius: 4,
        boxShadow: "0 4px 12px rgba(0,0,0,0.15)",
        zIndex: 20,
      }}
    >
      {options.length === 0 ? (
        <div style={{ padding: "6px 8px", fontSize: 12, color: "var(--secondary)" }}>(no options)</div>
      ) : (
        options.map((opt, i) => (
          <button
            key={`${opt}-${i}`}
            type="button"
            role="option"
            data-testid={`dv-option-${i}`}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => onPick(opt)}
            style={{
              display: "block",
              width: "100%",
              textAlign: "left",
              padding: "4px 8px",
              fontSize: 12,
              border: "none",
              background: "transparent",
              cursor: "pointer",
              color: "var(--foreground)",
            }}
            onMouseEnter={(e) => (e.currentTarget.style.background = "var(--hover)")}
            onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
          >
            {opt}
          </button>
        ))
      )}
    </div>
  );
}

// Re-export for convenience so XlsxEditor doesn't have to import from
// two adjacent modules.
export { singleSelection } from "./selection";
export type { CellPos, Selection } from "./selection";
