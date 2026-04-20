"use client";

import { useEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from "react";
import type { ChartKind, ChartPalette, Sheet, SheetChart } from "@officeai/xlsx";
import {
  bodyPxToAnchor,
  rawAnchorToBodyPx,
  type AnchorFromPx,
  type ResizeHandle,
} from "./ImageOverlay";
import type { AxisLookup } from "./gridDimensions";
import { parseChartRangeShape } from "./chartShape";

type AxisLike = ReadonlyArray<number> | AxisLookup;

interface ChartOverlayProps {
  readonly chart: SheetChart;
  readonly sheet: Sheet;
  readonly colXs: AxisLike;
  readonly rowYs: AxisLike;
  readonly headerOffset: { x: number; y: number };
  readonly selected?: boolean;
  readonly onSelect?: () => void;
  readonly onRequestRemove?: () => void;
  readonly onRequestEdit?: () => void;
  readonly onMoveCommit?: (anchor: AnchorFromPx) => void;
  readonly onResizeCommit?: (size: { widthPx: number; heightPx: number }) => void;
  readonly onChangeKind?: (kind: ChartKind) => void;
}

/**
 * Series color palettes. Keyed by {@link ChartPalette}; each is a
 * cycle the renderer walks via `colors[i % colors.length]`. Tweak
 * the hex values here and existing charts inherit the new look
 * because the model only stores the palette *name*.
 */
const PALETTES: Record<ChartPalette, ReadonlyArray<string>> = {
  default: ["#5b8def", "#f6c34a", "#7bc274", "#ec6f6f", "#9b6dd6", "#3eb6c4", "#f4a261", "#7d8eb1"],
  vibrant: ["#2563eb", "#dc2626", "#16a34a", "#f59e0b", "#9333ea", "#0891b2", "#db2777", "#65a30d"],
  pastel: ["#a8c5f0", "#fbd38d", "#9ed5a6", "#f5a8a8", "#c8a8e8", "#9bd3db", "#f5c79a", "#b8c1d6"],
  warm: ["#f97316", "#ef4444", "#fbbf24", "#dc2626", "#f59e0b", "#b91c1c", "#ea580c", "#facc15"],
  cool: ["#0ea5e9", "#10b981", "#6366f1", "#06b6d4", "#14b8a6", "#3b82f6", "#8b5cf6", "#22d3ee"],
  mono: ["#1f2937", "#374151", "#4b5563", "#6b7280", "#9ca3af", "#d1d5db", "#e5e7eb", "#f3f4f6"],
};

function paletteColors(palette: ChartPalette | undefined): ReadonlyArray<string> {
  return PALETTES[palette ?? "default"] ?? PALETTES.default;
}

/** Resolve renderer defaults so callers don't have to repeat them. */
function resolveStyle(chart: SheetChart): {
  readonly showLegend: boolean;
  readonly showDataLabels: boolean;
  readonly showGridlines: boolean;
  readonly xAxisTitle: string | undefined;
  readonly yAxisTitle: string | undefined;
} {
  return {
    showLegend: chart.showLegend !== false,
    showDataLabels: chart.showDataLabels === true,
    showGridlines: chart.showGridlines !== false,
    xAxisTitle: chart.xAxisTitle && chart.xAxisTitle.trim() !== "" ? chart.xAxisTitle : undefined,
    yAxisTitle: chart.yAxisTitle && chart.yAxisTitle.trim() !== "" ? chart.yAxisTitle : undefined,
  };
}

const HANDLE_SIZE = 9;
const TOOLBAR_HEIGHT = 28;
const TOOLBAR_GAP = 4;

const RESIZE_CURSORS: Record<ResizeHandle, string> = {
  nw: "nwse-resize",
  n: "ns-resize",
  ne: "nesw-resize",
  e: "ew-resize",
  se: "nwse-resize",
  s: "ns-resize",
  sw: "nesw-resize",
  w: "ew-resize",
};

const KIND_OPTIONS: ReadonlyArray<{ readonly kind: ChartKind; readonly label: string }> = [
  { kind: "column", label: "Col" },
  { kind: "bar", label: "Bar" },
  { kind: "line", label: "Line" },
  { kind: "pie", label: "Pie" },
];

type DragMode =
  | { kind: "move"; startMouseX: number; startMouseY: number; startBodyX: number; startBodyY: number }
  | {
      kind: "resize";
      handle: ResizeHandle;
      startMouseX: number;
      startMouseY: number;
      startBodyX: number;
      startBodyY: number;
      startWidth: number;
      startHeight: number;
    };

/**
 * Lightweight SVG renderer for an in-session {@link SheetChart}.
 *
 * Reads numeric cells live from {@link Sheet} so the chart updates
 * whenever the underlying data changes. Used by the Grid's chart
 * overlay layer; covers Column / Bar / Line / Pie. Not pixel-perfect
 * with Excel — close enough that the user can preview their data
 * shape before copying or sharing the workbook.
 *
 * Owns its own absolute positioning + pointer-driven drag/resize so
 * inline move and resize feel parity with image overlays. Geometry
 * commits to the parent on mouse-up via `onMoveCommit` /
 * `onResizeCommit` (which dispatch `xlsx:move-chart` /
 * `xlsx:resize-chart`) so each gesture is a single undo step. When
 * selected, a small toolbar floats above the chart with quick
 * chart-type switching, an "Edit data" entry that opens the edit
 * dialog, and a delete button.
 */
export function ChartOverlay(props: ChartOverlayProps): ReactNode {
  const {
    chart,
    sheet,
    colXs,
    rowYs,
    headerOffset,
    selected,
    onSelect,
    onRequestRemove,
    onRequestEdit,
    onMoveCommit,
    onResizeCommit,
    onChangeKind,
  } = props;
  const data = useMemo(() => extractSeries(chart, sheet), [chart, sheet]);

  const baseBodyPx = rawAnchorToBodyPx(chart.anchor, colXs, rowYs);
  const [dragMode, setDragMode] = useState<DragMode | null>(null);
  const [transient, setTransient] = useState<{
    bodyX: number;
    bodyY: number;
    widthPx: number;
    heightPx: number;
  } | null>(null);
  const transientRef = useRef(transient);
  useEffect(() => {
    transientRef.current = transient;
  }, [transient]);

  const displayBodyX = transient?.bodyX ?? baseBodyPx.x;
  const displayBodyY = transient?.bodyY ?? baseBodyPx.y;
  const displayWidth = transient?.widthPx ?? chart.anchor.widthPx;
  const displayHeight = transient?.heightPx ?? chart.anchor.heightPx;

  useEffect(() => {
    if (!dragMode) return;
    const onMove = (e: MouseEvent) => {
      if (dragMode.kind === "move") {
        const dx = e.clientX - dragMode.startMouseX;
        const dy = e.clientY - dragMode.startMouseY;
        setTransient({
          bodyX: Math.max(0, dragMode.startBodyX + dx),
          bodyY: Math.max(0, dragMode.startBodyY + dy),
          widthPx: chart.anchor.widthPx,
          heightPx: chart.anchor.heightPx,
        });
        return;
      }
      const dx = e.clientX - dragMode.startMouseX;
      const dy = e.clientY - dragMode.startMouseY;
      setTransient(applyResize(dragMode, dx, dy));
    };
    const onUp = () => {
      const final = transientRef.current;
      const mode = dragMode;
      setDragMode(null);
      setTransient(null);
      if (!final) return;
      if (mode.kind === "move") {
        const anchor = bodyPxToAnchor(final.bodyX, final.bodyY, colXs, rowYs);
        if (anchorChanged(anchor, chart.anchor)) onMoveCommit?.(anchor);
        return;
      }
      const widthPx = Math.max(80, Math.round(final.widthPx));
      const heightPx = Math.max(60, Math.round(final.heightPx));
      if (widthPx !== chart.anchor.widthPx || heightPx !== chart.anchor.heightPx) {
        onResizeCommit?.({ widthPx, heightPx });
      }
      if (mode.handle.includes("w") || mode.handle.includes("n")) {
        const anchor = bodyPxToAnchor(final.bodyX, final.bodyY, colXs, rowYs);
        if (anchorChanged(anchor, chart.anchor)) onMoveCommit?.(anchor);
      }
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
  }, [dragMode, colXs, rowYs, chart.anchor, onMoveCommit, onResizeCommit]);

  const titleHeight = chart.title ? 22 : 0;
  const innerHeight = Math.max(40, displayHeight - titleHeight - 8);
  const innerWidth = Math.max(40, displayWidth - 8);

  const wrapperStyle: CSSProperties = {
    position: "absolute",
    top: headerOffset.y + displayBodyY,
    left: headerOffset.x + displayBodyX,
    width: displayWidth,
    height: displayHeight,
    zIndex: selected ? 12 : 5,
    userSelect: "none",
  };

  const bodyStyle: CSSProperties = {
    width: "100%",
    height: "100%",
    background: "var(--background)",
    border: selected ? "2px solid var(--ai-violet)" : "1px solid var(--border)",
    borderRadius: 6,
    boxShadow: selected ? "0 4px 12px rgba(0,0,0,0.08)" : "0 1px 3px rgba(0,0,0,0.05)",
    boxSizing: "border-box",
    overflow: "hidden",
    cursor: dragMode?.kind === "move" ? "grabbing" : selected ? "grab" : "default",
    display: "flex",
    flexDirection: "column",
  };

  const startMove = (e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    onSelect?.();
    if (!onMoveCommit) return;
    setTransient({
      bodyX: baseBodyPx.x,
      bodyY: baseBodyPx.y,
      widthPx: chart.anchor.widthPx,
      heightPx: chart.anchor.heightPx,
    });
    setDragMode({
      kind: "move",
      startMouseX: e.clientX,
      startMouseY: e.clientY,
      startBodyX: baseBodyPx.x,
      startBodyY: baseBodyPx.y,
    });
  };

  return (
    <div
      data-testid={`chart-overlay-${chart.id}`}
      data-chart-id={chart.id}
      style={wrapperStyle}
    >
      {selected ? (
        <ChartToolbar
          chart={chart}
          onChangeKind={onChangeKind}
          onRequestEdit={onRequestEdit}
          onRequestRemove={onRequestRemove}
        />
      ) : null}
      <div
        role="figure"
        aria-label={chart.title ?? `${chart.kind} chart`}
        style={bodyStyle}
        onMouseDown={startMove}
      >
        {chart.title ? (
          <div
            style={{
              height: titleHeight,
              padding: "0 8px",
              fontSize: 12,
              fontWeight: 600,
              color: "var(--foreground)",
              display: "flex",
              alignItems: "center",
            }}
          >
            <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              {chart.title}
            </span>
          </div>
        ) : null}
        <div style={{ flex: 1, padding: 4, pointerEvents: "none" }}>
          <ChartCanvas
            kind={chart.kind}
            width={innerWidth}
            height={innerHeight}
            categories={data.categories}
            series={data.series}
            style={resolveStyle(chart)}
            paletteColors={paletteColors(chart.palette)}
          />
        </div>
      </div>
      {selected
        ? (["nw", "n", "ne", "e", "se", "s", "sw", "w"] as ResizeHandle[]).map((h) => (
            <div
              key={`chart-handle-${h}`}
              data-testid={`chart-handle-${chart.id}-${h}`}
              aria-label={`Resize ${h}`}
              onMouseDown={(e) => {
                e.preventDefault();
                e.stopPropagation();
                onSelect?.();
                if (!onResizeCommit) return;
                setTransient({
                  bodyX: baseBodyPx.x,
                  bodyY: baseBodyPx.y,
                  widthPx: chart.anchor.widthPx,
                  heightPx: chart.anchor.heightPx,
                });
                setDragMode({
                  kind: "resize",
                  handle: h,
                  startMouseX: e.clientX,
                  startMouseY: e.clientY,
                  startBodyX: baseBodyPx.x,
                  startBodyY: baseBodyPx.y,
                  startWidth: chart.anchor.widthPx,
                  startHeight: chart.anchor.heightPx,
                });
              }}
              style={{
                position: "absolute",
                ...handleStyle(h),
                width: HANDLE_SIZE,
                height: HANDLE_SIZE,
                background: "white",
                border: "1.5px solid var(--ai-violet, #7c3aed)",
                borderRadius: 2,
                cursor: RESIZE_CURSORS[h],
                zIndex: 13,
              }}
            />
          ))
        : null}
    </div>
  );
}

interface ChartToolbarProps {
  readonly chart: SheetChart;
  readonly onChangeKind?: (kind: ChartKind) => void;
  readonly onRequestEdit?: () => void;
  readonly onRequestRemove?: () => void;
}

/**
 * Floating chart toolbar shown when the chart is selected. Sits a
 * few pixels above the chart body so it doesn't shrink the chart
 * itself. Quick chart-type switcher + Edit + Delete; mousedown is
 * stopped so clicks here don't accidentally start a drag on the
 * underlying chart body.
 */
function ChartToolbar(props: ChartToolbarProps): ReactNode {
  const { chart, onChangeKind, onRequestEdit, onRequestRemove } = props;
  return (
    <div
      data-testid={`chart-toolbar-${chart.id}`}
      onMouseDown={(e) => {
        e.preventDefault();
        e.stopPropagation();
      }}
      style={{
        position: "absolute",
        top: -(TOOLBAR_HEIGHT + TOOLBAR_GAP),
        left: 0,
        height: TOOLBAR_HEIGHT,
        display: "flex",
        alignItems: "center",
        gap: 4,
        padding: "0 6px",
        background: "var(--background)",
        border: "1px solid var(--border)",
        borderRadius: 6,
        boxShadow: "0 4px 12px rgba(0,0,0,0.12)",
        fontSize: 11,
        zIndex: 14,
      }}
    >
      {KIND_OPTIONS.map((opt) => {
        const active = opt.kind === chart.kind;
        return (
          <button
            key={opt.kind}
            type="button"
            data-testid={`chart-toolbar-kind-${opt.kind}`}
            disabled={!onChangeKind}
            onClick={() => onChangeKind?.(opt.kind)}
            title={`${opt.label} chart`}
            style={{
              padding: "2px 8px",
              border: active ? "1px solid var(--ai-violet)" : "1px solid var(--border)",
              borderRadius: 4,
              background: active ? "var(--ai-violet-light, rgba(124,58,237,0.12))" : "var(--background)",
              color: "var(--foreground)",
              fontSize: 11,
              cursor: onChangeKind ? "pointer" : "default",
            }}
          >
            {opt.label}
          </button>
        );
      })}
      <div style={{ width: 1, height: 16, background: "var(--divider)" }} aria-hidden />
      {onRequestEdit ? (
        <button
          type="button"
          data-testid="chart-toolbar-edit"
          onClick={onRequestEdit}
          title="Edit chart data"
          style={{
            padding: "2px 8px",
            border: "1px solid var(--border)",
            borderRadius: 4,
            background: "var(--background)",
            color: "var(--foreground)",
            fontSize: 11,
            cursor: "pointer",
          }}
        >
          Edit data…
        </button>
      ) : null}
      {onRequestRemove ? (
        <button
          type="button"
          data-testid="chart-toolbar-remove"
          aria-label="Remove chart"
          title="Remove chart"
          onClick={onRequestRemove}
          style={{
            padding: "1px 6px",
            border: "1px solid var(--border)",
            borderRadius: 4,
            background: "transparent",
            color: "var(--muted-foreground)",
            fontSize: 13,
            cursor: "pointer",
          }}
        >
          ×
        </button>
      ) : null}
    </div>
  );
}

function applyResize(
  drag: Extract<DragMode, { kind: "resize" }>,
  dx: number,
  dy: number
): { bodyX: number; bodyY: number; widthPx: number; heightPx: number } {
  let { startBodyX: x, startBodyY: y, startWidth: w, startHeight: h } = drag;
  const minW = 80;
  const minH = 60;
  if (drag.handle.includes("e")) w = Math.max(minW, drag.startWidth + dx);
  if (drag.handle.includes("s")) h = Math.max(minH, drag.startHeight + dy);
  if (drag.handle.includes("w")) {
    const newW = Math.max(minW, drag.startWidth - dx);
    x = drag.startBodyX + (drag.startWidth - newW);
    w = newW;
  }
  if (drag.handle.includes("n")) {
    const newH = Math.max(minH, drag.startHeight - dy);
    y = drag.startBodyY + (drag.startHeight - newH);
    h = newH;
  }
  return { bodyX: Math.max(0, x), bodyY: Math.max(0, y), widthPx: w, heightPx: h };
}

function handleStyle(h: ResizeHandle): CSSProperties {
  const half = -Math.floor(HANDLE_SIZE / 2);
  switch (h) {
    case "nw":
      return { top: half, left: half };
    case "n":
      return { top: half, left: `calc(50% - ${HANDLE_SIZE / 2}px)` };
    case "ne":
      return { top: half, right: half };
    case "e":
      return { top: `calc(50% - ${HANDLE_SIZE / 2}px)`, right: half };
    case "se":
      return { bottom: half, right: half };
    case "s":
      return { bottom: half, left: `calc(50% - ${HANDLE_SIZE / 2}px)` };
    case "sw":
      return { bottom: half, left: half };
    case "w":
      return { top: `calc(50% - ${HANDLE_SIZE / 2}px)`, left: half };
  }
}

function anchorChanged(next: AnchorFromPx, current: SheetChart["anchor"]): boolean {
  return (
    next.fromRow !== current.fromRow ||
    next.fromCol !== current.fromCol ||
    Math.round(next.fromOffsetXPx) !== Math.round(current.fromOffsetXPx) ||
    Math.round(next.fromOffsetYPx) !== Math.round(current.fromOffsetYPx)
  );
}

interface ChartSeries {
  readonly name: string;
  readonly values: ReadonlyArray<number>;
  readonly color: string;
}

interface ChartData {
  readonly categories: ReadonlyArray<string>;
  readonly series: ReadonlyArray<ChartSeries>;
}

function extractSeries(chart: SheetChart, sheet: Sheet): ChartData {
  const range = parseChartRangeShape(chart.dataRange);
  if (!range) return { categories: [], series: [] };
  const { r1, r2, c1, c2 } = range;
  const headerRow = chart.hasHeaderRow ? r1 : -1;
  const bodyStart = chart.hasHeaderRow ? r1 + 1 : r1;
  const catCol = chart.hasCategoryColumn ? c1 : -1;
  const valStart = chart.hasCategoryColumn ? c1 + 1 : c1;
  const colors = paletteColors(chart.palette);

  const categories: string[] = [];
  for (let r = bodyStart; r <= r2; r++) {
    if (catCol === -1) {
      categories.push(String(r - bodyStart + 1));
      continue;
    }
    const cell = sheet.cells.get(`${r}:${catCol}`);
    categories.push(formatCategory(cell?.value));
  }

  const series: ChartSeries[] = [];
  for (let c = valStart; c <= c2; c++) {
    const headerCell = headerRow === -1 ? undefined : sheet.cells.get(`${headerRow}:${c}`);
    const name = headerCell ? formatCategory(headerCell.value) : `Series ${c - valStart + 1}`;
    const values: number[] = [];
    for (let r = bodyStart; r <= r2; r++) {
      const cell = sheet.cells.get(`${r}:${c}`);
      const n = numericValue(cell?.value);
      values.push(n);
    }
    const color = colors[(c - valStart) % colors.length]!;
    series.push({ name, values, color });
  }
  return { categories, series };
}

function formatCategory(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (typeof value === "object" && value && "kind" in value && (value as { kind: string }).kind === "error") {
    return "";
  }
  return "";
}

function numericValue(value: unknown): number {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (typeof value === "string") {
    const n = Number.parseFloat(value);
    if (Number.isFinite(n)) return n;
  }
  return 0;
}

interface ChartStyle {
  readonly showLegend: boolean;
  readonly showDataLabels: boolean;
  readonly showGridlines: boolean;
  readonly xAxisTitle: string | undefined;
  readonly yAxisTitle: string | undefined;
}

interface ChartCanvasProps {
  readonly kind: ChartKind;
  readonly width: number;
  readonly height: number;
  readonly categories: ReadonlyArray<string>;
  readonly series: ReadonlyArray<ChartSeries>;
  readonly style: ChartStyle;
  readonly paletteColors: ReadonlyArray<string>;
}

function ChartCanvas(props: ChartCanvasProps): ReactNode {
  const { kind, width, height, categories, series, style, paletteColors: colors } = props;
  if (series.length === 0 || categories.length === 0) {
    return (
      <div
        style={{
          width,
          height,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          fontSize: 11,
          color: "var(--muted-foreground)",
        }}
      >
        No data in selected range
      </div>
    );
  }
  switch (kind) {
    case "column":
      return (
        <ColumnOrBar
          width={width}
          height={height}
          categories={categories}
          series={series}
          style={style}
          horizontal={false}
        />
      );
    case "bar":
      return (
        <ColumnOrBar
          width={width}
          height={height}
          categories={categories}
          series={series}
          style={style}
          horizontal={true}
        />
      );
    case "line":
      return (
        <LineChart
          width={width}
          height={height}
          categories={categories}
          series={series}
          style={style}
        />
      );
    case "pie":
      return (
        <PieChart
          width={width}
          height={height}
          series={series[0]!}
          categories={categories}
          style={style}
          paletteColors={colors}
        />
      );
    default: {
      const _exhaustive: never = kind;
      void _exhaustive;
      return null;
    }
  }
}

interface BaseChartProps {
  readonly width: number;
  readonly height: number;
  readonly categories: ReadonlyArray<string>;
  readonly series: ReadonlyArray<ChartSeries>;
  readonly style: ChartStyle;
}

function ColumnOrBar(props: BaseChartProps & { horizontal: boolean }): ReactNode {
  const { width, height, categories, series, style, horizontal } = props;
  // Pad the axis side that gets the title so glyphs don't overlap the
  // tick labels. We also reserve a bit of space at the top when a
  // legend is shown so it doesn't crash into the plot area.
  const legendPad = style.showLegend && series.length > 1 ? 12 : 0;
  const padding = horizontal
    ? {
        top: 8 + legendPad,
        right: 12,
        bottom: style.xAxisTitle ? 40 : 24,
        left: style.yAxisTitle ? 52 : 36,
      }
    : {
        top: 8 + legendPad,
        right: 12,
        bottom: style.xAxisTitle ? 40 : 24,
        left: style.yAxisTitle ? 52 : 36,
      };
  const plotW = Math.max(10, width - padding.left - padding.right);
  const plotH = Math.max(10, height - padding.top - padding.bottom);
  const groupCount = categories.length;
  const seriesCount = series.length;
  const allValues = series.flatMap((s) => s.values);
  const minRaw = Math.min(0, ...allValues);
  const maxRaw = Math.max(0, ...allValues);
  const max = maxRaw === minRaw ? maxRaw + 1 : maxRaw;
  const min = minRaw;
  const span = max - min || 1;

  const ticks = niceTicks(min, max, 4);
  const xs = (i: number) => padding.left + (i + 0.5) * (plotW / groupCount);
  const ys = (v: number) => padding.top + (1 - (v - min) / span) * plotH;
  const groupWidth = plotW / groupCount;
  const barWidth = Math.max(2, (groupWidth * 0.8) / seriesCount);

  if (horizontal) {
    const bandH = plotH / groupCount;
    const barH = Math.max(2, (bandH * 0.8) / seriesCount);
    const xv = (v: number) => padding.left + ((v - min) / span) * plotW;
    return (
      <svg width={width} height={height}>
        {ticks.map((t) => (
          <g key={`gx-${t}`}>
            {style.showGridlines ? (
              <line
                x1={xv(t)}
                x2={xv(t)}
                y1={padding.top}
                y2={padding.top + plotH}
                stroke="var(--divider)"
              />
            ) : null}
            <text
              x={xv(t)}
              y={padding.top + plotH + 12}
              fontSize={9}
              textAnchor="middle"
              fill="var(--muted-foreground)"
            >
              {formatTick(t)}
            </text>
          </g>
        ))}
        {categories.map((c, i) => (
          <text
            key={`cat-${i}`}
            x={padding.left - 4}
            y={padding.top + (i + 0.5) * bandH + 3}
            fontSize={9}
            textAnchor="end"
            fill="var(--muted-foreground)"
          >
            {truncate(c, 10)}
          </text>
        ))}
        {series.map((s, si) =>
          s.values.map((v, i) => {
            const groupTop = padding.top + i * bandH;
            const top = groupTop + bandH * 0.1 + si * barH;
            const x0 = xv(0);
            const x1 = xv(v);
            const x = Math.min(x0, x1);
            const w = Math.abs(x1 - x0);
            return (
              <g key={`bar-${si}-${i}`}>
                <rect x={x} y={top} width={w} height={barH} fill={s.color} opacity={0.9} />
                {style.showDataLabels ? (
                  <text
                    x={v >= 0 ? x + w + 3 : x - 3}
                    y={top + barH / 2 + 3}
                    fontSize={9}
                    textAnchor={v >= 0 ? "start" : "end"}
                    fill="var(--foreground)"
                  >
                    {formatLabel(v)}
                  </text>
                ) : null}
              </g>
            );
          })
        )}
        {style.xAxisTitle ? (
          <text
            x={padding.left + plotW / 2}
            y={height - 6}
            fontSize={10}
            fontWeight={500}
            textAnchor="middle"
            fill="var(--foreground)"
          >
            {truncate(style.xAxisTitle, 60)}
          </text>
        ) : null}
        {style.yAxisTitle ? (
          <text
            x={12}
            y={padding.top + plotH / 2}
            fontSize={10}
            fontWeight={500}
            textAnchor="middle"
            fill="var(--foreground)"
            transform={`rotate(-90, 12, ${padding.top + plotH / 2})`}
          >
            {truncate(style.yAxisTitle, 60)}
          </text>
        ) : null}
        {style.showLegend ? <Legend series={series} x={width - 8} y={2} anchor="end" /> : null}
      </svg>
    );
  }

  return (
    <svg width={width} height={height}>
      {ticks.map((t) => (
        <g key={`gy-${t}`}>
          {style.showGridlines ? (
            <line
              x1={padding.left}
              x2={padding.left + plotW}
              y1={ys(t)}
              y2={ys(t)}
              stroke="var(--divider)"
            />
          ) : null}
          <text
            x={padding.left - 4}
            y={ys(t) + 3}
            fontSize={9}
            textAnchor="end"
            fill="var(--muted-foreground)"
          >
            {formatTick(t)}
          </text>
        </g>
      ))}
      {categories.map((c, i) => (
        <text
          key={`cat-${i}`}
          x={xs(i)}
          y={padding.top + plotH + 12}
          fontSize={9}
          textAnchor="middle"
          fill="var(--muted-foreground)"
        >
          {truncate(c, 8)}
        </text>
      ))}
      {series.map((s, si) =>
        s.values.map((v, i) => {
          const groupLeft = padding.left + i * groupWidth;
          const left = groupLeft + groupWidth * 0.1 + si * barWidth;
          const y0 = ys(0);
          const y1 = ys(v);
          const top = Math.min(y0, y1);
          const h = Math.abs(y1 - y0);
          return (
            <g key={`col-${si}-${i}`}>
              <rect x={left} y={top} width={barWidth} height={h} fill={s.color} opacity={0.9} />
              {style.showDataLabels ? (
                <text
                  x={left + barWidth / 2}
                  y={v >= 0 ? top - 3 : top + h + 9}
                  fontSize={9}
                  textAnchor="middle"
                  fill="var(--foreground)"
                >
                  {formatLabel(v)}
                </text>
              ) : null}
            </g>
          );
        })
      )}
      {style.xAxisTitle ? (
        <text
          x={padding.left + plotW / 2}
          y={height - 6}
          fontSize={10}
          fontWeight={500}
          textAnchor="middle"
          fill="var(--foreground)"
        >
          {truncate(style.xAxisTitle, 60)}
        </text>
      ) : null}
      {style.yAxisTitle ? (
        <text
          x={12}
          y={padding.top + plotH / 2}
          fontSize={10}
          fontWeight={500}
          textAnchor="middle"
          fill="var(--foreground)"
          transform={`rotate(-90, 12, ${padding.top + plotH / 2})`}
        >
          {truncate(style.yAxisTitle, 60)}
        </text>
      ) : null}
      {style.showLegend ? <Legend series={series} x={width - 8} y={2} anchor="end" /> : null}
    </svg>
  );
}

function LineChart(props: BaseChartProps): ReactNode {
  const { width, height, categories, series, style } = props;
  const legendPad = style.showLegend && series.length > 1 ? 12 : 0;
  const padding = {
    top: 8 + legendPad,
    right: 12,
    bottom: style.xAxisTitle ? 40 : 24,
    left: style.yAxisTitle ? 52 : 36,
  };
  const plotW = Math.max(10, width - padding.left - padding.right);
  const plotH = Math.max(10, height - padding.top - padding.bottom);
  const allValues = series.flatMap((s) => s.values);
  const minRaw = Math.min(...allValues);
  const maxRaw = Math.max(...allValues);
  const min = minRaw === maxRaw ? minRaw - 1 : minRaw;
  const max = minRaw === maxRaw ? maxRaw + 1 : maxRaw;
  const span = max - min || 1;
  const xs = (i: number) =>
    padding.left + (categories.length === 1 ? plotW / 2 : (i / (categories.length - 1)) * plotW);
  const ys = (v: number) => padding.top + (1 - (v - min) / span) * plotH;
  const ticks = niceTicks(min, max, 4);
  return (
    <svg width={width} height={height}>
      {ticks.map((t) => (
        <g key={`gy-${t}`}>
          {style.showGridlines ? (
            <line
              x1={padding.left}
              x2={padding.left + plotW}
              y1={ys(t)}
              y2={ys(t)}
              stroke="var(--divider)"
            />
          ) : null}
          <text
            x={padding.left - 4}
            y={ys(t) + 3}
            fontSize={9}
            textAnchor="end"
            fill="var(--muted-foreground)"
          >
            {formatTick(t)}
          </text>
        </g>
      ))}
      {categories.map((c, i) => (
        <text
          key={`cat-${i}`}
          x={xs(i)}
          y={padding.top + plotH + 12}
          fontSize={9}
          textAnchor="middle"
          fill="var(--muted-foreground)"
        >
          {truncate(c, 8)}
        </text>
      ))}
      {series.map((s, si) => {
        const path = s.values.map((v, i) => `${i === 0 ? "M" : "L"} ${xs(i)} ${ys(v)}`).join(" ");
        return (
          <g key={`line-${si}`}>
            <path d={path} stroke={s.color} strokeWidth={1.8} fill="none" />
            {s.values.map((v, i) => (
              <circle key={`pt-${si}-${i}`} cx={xs(i)} cy={ys(v)} r={2.5} fill={s.color} />
            ))}
            {style.showDataLabels
              ? s.values.map((v, i) => (
                  <text
                    key={`lbl-${si}-${i}`}
                    x={xs(i)}
                    y={ys(v) - 6}
                    fontSize={9}
                    textAnchor="middle"
                    fill="var(--foreground)"
                  >
                    {formatLabel(v)}
                  </text>
                ))
              : null}
          </g>
        );
      })}
      {style.xAxisTitle ? (
        <text
          x={padding.left + plotW / 2}
          y={height - 6}
          fontSize={10}
          fontWeight={500}
          textAnchor="middle"
          fill="var(--foreground)"
        >
          {truncate(style.xAxisTitle, 60)}
        </text>
      ) : null}
      {style.yAxisTitle ? (
        <text
          x={12}
          y={padding.top + plotH / 2}
          fontSize={10}
          fontWeight={500}
          textAnchor="middle"
          fill="var(--foreground)"
          transform={`rotate(-90, 12, ${padding.top + plotH / 2})`}
        >
          {truncate(style.yAxisTitle, 60)}
        </text>
      ) : null}
      {style.showLegend ? <Legend series={series} x={width - 8} y={2} anchor="end" /> : null}
    </svg>
  );
}

function PieChart(props: {
  readonly width: number;
  readonly height: number;
  readonly series: ChartSeries;
  readonly categories: ReadonlyArray<string>;
  readonly style: ChartStyle;
  readonly paletteColors: ReadonlyArray<string>;
}): ReactNode {
  const { width, height, series, categories, style, paletteColors: colors } = props;
  const cx = width / 2;
  const cy = height / 2 + 4;
  const radius = Math.max(10, Math.min(width, height) / 2 - 16);
  const total = series.values.reduce((acc, v) => acc + Math.max(0, v), 0);
  if (total <= 0) {
    return (
      <div
        style={{
          width,
          height,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          fontSize: 11,
          color: "var(--muted-foreground)",
        }}
      >
        Pie needs positive values
      </div>
    );
  }
  // Pre-compute slice geometry so we can paint slices and label
  // overlays in two separate passes (labels must sit above every
  // slice, not just the one immediately under them).
  let acc = -Math.PI / 2;
  const slices = series.values.map((v, i) => {
    const value = Math.max(0, v);
    const angle = (value / total) * Math.PI * 2;
    const start = acc;
    const end = acc + angle;
    acc = end;
    return { value, angle, start, end, color: colors[i % colors.length]! };
  });
  return (
    <svg width={width} height={height}>
      {slices.map((s, i) => {
        const x0 = cx + radius * Math.cos(s.start);
        const y0 = cy + radius * Math.sin(s.start);
        const x1 = cx + radius * Math.cos(s.end);
        const y1 = cy + radius * Math.sin(s.end);
        const large = s.angle > Math.PI ? 1 : 0;
        return (
          <path
            key={`slice-${i}`}
            d={`M ${cx} ${cy} L ${x0} ${y0} A ${radius} ${radius} 0 ${large} 1 ${x1} ${y1} Z`}
            fill={s.color}
            opacity={0.9}
          />
        );
      })}
      {style.showDataLabels
        ? slices.map((s, i) => {
            if (s.angle <= 0) return null;
            const mid = (s.start + s.end) / 2;
            const lx = cx + radius * 0.62 * Math.cos(mid);
            const ly = cy + radius * 0.62 * Math.sin(mid);
            const pct = total > 0 ? Math.round((s.value / total) * 100) : 0;
            return (
              <text
                key={`pielbl-${i}`}
                x={lx}
                y={ly + 3}
                fontSize={9}
                textAnchor="middle"
                fill="white"
                style={{ paintOrder: "stroke", stroke: "rgba(0,0,0,0.35)", strokeWidth: 2 }}
              >
                {`${pct}%`}
              </text>
            );
          })
        : null}
      {style.showLegend
        ? categories.map((c, i) => (
            <g key={`legend-${i}`} transform={`translate(${4}, ${10 + i * 12})`}>
              <rect width={8} height={8} fill={colors[i % colors.length]!} />
              <text x={12} y={7} fontSize={9} fill="var(--muted-foreground)">
                {truncate(c, 14)}
              </text>
            </g>
          ))
        : null}
    </svg>
  );
}

function Legend(props: {
  readonly series: ReadonlyArray<ChartSeries>;
  readonly x: number;
  readonly y: number;
  readonly anchor: "start" | "end";
}): ReactNode {
  const { series, x, y, anchor } = props;
  if (series.length === 1) return null;
  return (
    <g transform={`translate(${x}, ${y})`} textAnchor={anchor}>
      {series.map((s, i) => (
        <g key={`legend-${i}`} transform={`translate(0, ${i * 12})`}>
          <text x={anchor === "end" ? -12 : 12} y={8} fontSize={9} fill="var(--muted-foreground)">
            {truncate(s.name, 12)}
          </text>
          <rect x={anchor === "end" ? -8 : 0} y={2} width={8} height={8} fill={s.color} />
        </g>
      ))}
    </g>
  );
}

function niceTicks(min: number, max: number, count: number): number[] {
  if (!Number.isFinite(min) || !Number.isFinite(max) || min === max) return [min];
  const range = niceNum(max - min, false);
  const step = niceNum(range / Math.max(1, count - 1), true);
  const niceMin = Math.floor(min / step) * step;
  const niceMax = Math.ceil(max / step) * step;
  const out: number[] = [];
  for (let v = niceMin; v <= niceMax + step / 2; v += step) {
    out.push(Number.parseFloat(v.toFixed(10)));
  }
  return out;
}

function niceNum(range: number, round: boolean): number {
  const exp = Math.floor(Math.log10(Math.abs(range) || 1));
  const fraction = (Math.abs(range) || 1) / Math.pow(10, exp);
  let nice: number;
  if (round) {
    if (fraction < 1.5) nice = 1;
    else if (fraction < 3) nice = 2;
    else if (fraction < 7) nice = 5;
    else nice = 10;
  } else {
    if (fraction <= 1) nice = 1;
    else if (fraction <= 2) nice = 2;
    else if (fraction <= 5) nice = 5;
    else nice = 10;
  }
  return nice * Math.pow(10, exp);
}

function formatTick(value: number): string {
  if (Math.abs(value) >= 1000) return `${(value / 1000).toFixed(1)}k`;
  return Number.isInteger(value) ? String(value) : value.toFixed(1);
}

/**
 * Format a single data value for an inline label. Same shorthand
 * as {@link formatTick} for big numbers; never widens the chart by
 * more than a handful of glyphs even for awkward decimals.
 */
function formatLabel(value: number): string {
  if (!Number.isFinite(value)) return "";
  if (Math.abs(value) >= 1000) return `${(value / 1000).toFixed(1)}k`;
  if (Number.isInteger(value)) return String(value);
  return value.toFixed(2);
}

function truncate(s: string, max: number): string {
  if (s.length <= max) return s;
  return `${s.slice(0, Math.max(1, max - 1))}…`;
}
