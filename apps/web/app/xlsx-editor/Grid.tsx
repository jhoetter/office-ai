"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { CSSProperties, ReactNode } from "react";
import { cellKey, colToLetter, type Sheet, type CellValue } from "@officeai/xlsx";

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

export interface GridSelection {
  readonly row: number;
  readonly col: number;
}

export interface GridProps {
  readonly sheet: Sheet;
  readonly selection: GridSelection | null;
  readonly onSelect: (sel: GridSelection) => void;
  /**
   * Called when the user commits an in-cell edit (double-click → type
   * → Enter). Parent dispatches `xlsx:set-cell-value` /
   * `xlsx:set-cell-formula`.
   */
  readonly onCommitEdit: (sel: GridSelection, value: string) => void;
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
 */
export function Grid(props: GridProps): ReactNode {
  const { sheet, selection, onSelect, onCommitEdit } = props;

  const scrollRef = useRef<HTMLDivElement | null>(null);
  const [scroll, setScroll] = useState({ top: 0, left: 0 });
  const [editing, setEditing] = useState<{ row: number; col: number; draft: string } | null>(null);
  const [viewport, setViewport] = useState({ width: 0, height: 0 });

  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const onResize = () => setViewport({ width: el.clientWidth, height: el.clientHeight });
    onResize();
    const ro = new ResizeObserver(onResize);
    ro.observe(el);
    return () => ro.disconnect();
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

  const cellList: ReactNode[] = useMemo(() => {
    const out: ReactNode[] = [];
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const cell = sheet.cells.get(cellKey(r, c));
        const display = cell ? formatDisplay(cell.value) : "";
        const isSelected = selection?.row === r && selection?.col === c;
        const isEditing = editing?.row === r && editing?.col === c;
        out.push(
          <div
            key={`c-${r}-${c}`}
            data-testid={`cell-${colToLetter(c)}${r + 1}`}
            role="gridcell"
            aria-selected={isSelected || undefined}
            onMouseDown={() => {
              if (!isEditing) onSelect({ row: r, col: c });
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
              width: COL_WIDTH,
              height: ROW_HEIGHT,
              boxSizing: "border-box",
              borderRight: "1px solid var(--divider)",
              borderBottom: "1px solid var(--divider)",
              padding: "0 6px",
              display: "flex",
              alignItems: "center",
              fontSize: 12,
              lineHeight: `${ROW_HEIGHT - 2}px`,
              background: isSelected ? "var(--ai-violet-light)" : "var(--background)",
              outline: isSelected ? "2px solid var(--ai-violet)" : "none",
              outlineOffset: -2,
              color: "var(--foreground)",
              overflow: "hidden",
              whiteSpace: "nowrap",
              textOverflow: "ellipsis",
              cursor: "cell",
              userSelect: "none",
              zIndex: isSelected ? 1 : 0,
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
  }, [sheet, startRow, endRow, startCol, endCol, selection, editing, onSelect, onCommitEdit]);

  // Column header band — tracks the viewport's top edge.
  const colHeaders: ReactNode[] = [];
  for (let c = startCol; c <= endCol; c++) {
    const isActive = selection?.col === c;
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
    const isActive = selection?.row === r;
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
      </div>
    </div>
  );
}

function formatDisplay(value: CellValue): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  switch (value.kind) {
    case "error":
      return value.code;
    default: {
      const _exhaustive: never = value.kind;
      void _exhaustive;
      return "";
    }
  }
}
