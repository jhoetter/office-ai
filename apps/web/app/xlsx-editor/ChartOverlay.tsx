"use client";

import { useMemo, type CSSProperties, type ReactNode } from "react";
import type { ChartKind, Sheet, SheetChart } from "@officeai/xlsx";

interface ChartOverlayProps {
  readonly chart: SheetChart;
  readonly sheet: Sheet;
  readonly width: number;
  readonly height: number;
  readonly selected?: boolean;
  readonly onSelect?: () => void;
  readonly onRequestRemove?: () => void;
}

const PALETTE = ["#5b8def", "#f6c34a", "#7bc274", "#ec6f6f", "#9b6dd6", "#3eb6c4", "#f4a261", "#7d8eb1"];

/**
 * Lightweight SVG renderer for an in-session {@link SheetChart}.
 *
 * Reads numeric cells live from {@link Sheet} so the chart updates
 * whenever the underlying data changes. Used by the Grid's chart
 * overlay layer; covers Column / Bar / Line / Pie. Not pixel-perfect
 * with Excel — close enough that the user can preview their data
 * shape before copying or sharing the workbook.
 */
export function ChartOverlay(props: ChartOverlayProps): ReactNode {
  const { chart, sheet, width, height, selected, onSelect, onRequestRemove } = props;
  const data = useMemo(() => extractSeries(chart, sheet), [chart, sheet]);

  const containerStyle: CSSProperties = {
    width,
    height,
    background: "var(--background)",
    border: selected ? "2px solid var(--ai-violet)" : "1px solid var(--border)",
    borderRadius: 6,
    boxShadow: selected ? "0 4px 12px rgba(0,0,0,0.08)" : "0 1px 3px rgba(0,0,0,0.05)",
    boxSizing: "border-box",
    overflow: "hidden",
    cursor: "default",
    display: "flex",
    flexDirection: "column",
  };

  const titleHeight = chart.title ? 22 : 0;
  const innerHeight = Math.max(40, height - titleHeight - 8);
  const innerWidth = Math.max(40, width - 8);

  return (
    <div
      data-testid={`chart-overlay-${chart.id}`}
      role="figure"
      aria-label={chart.title ?? `${chart.kind} chart`}
      style={containerStyle}
      onMouseDown={(e) => {
        e.stopPropagation();
        onSelect?.();
      }}
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
            justifyContent: "space-between",
          }}
        >
          <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
            {chart.title}
          </span>
          {selected && onRequestRemove ? (
            <button
              type="button"
              onClick={(e) => {
                e.stopPropagation();
                onRequestRemove();
              }}
              style={{
                fontSize: 11,
                background: "transparent",
                border: "1px solid var(--border)",
                borderRadius: 4,
                padding: "1px 6px",
                cursor: "pointer",
                color: "var(--muted-foreground)",
              }}
              aria-label="Remove chart"
            >
              ×
            </button>
          ) : null}
        </div>
      ) : null}
      <div style={{ flex: 1, padding: 4 }}>
        <ChartCanvas
          kind={chart.kind}
          width={innerWidth}
          height={innerHeight}
          categories={data.categories}
          series={data.series}
        />
      </div>
    </div>
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
  const range = parseA1Range(chart.dataRange);
  if (!range) return { categories: [], series: [] };
  const { r1, r2, c1, c2 } = range;
  const headerRow = chart.hasHeaderRow ? r1 : -1;
  const bodyStart = chart.hasHeaderRow ? r1 + 1 : r1;
  const catCol = chart.hasCategoryColumn ? c1 : -1;
  const valStart = chart.hasCategoryColumn ? c1 + 1 : c1;

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
    const color = PALETTE[(c - valStart) % PALETTE.length]!;
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

function parseA1Range(range: string): { r1: number; r2: number; c1: number; c2: number } | null {
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(range);
  if (!m) return null;
  const c1 = letterToCol(m[1]!);
  const r1 = Number.parseInt(m[2]!, 10) - 1;
  const c2 = letterToCol(m[3]!);
  const r2 = Number.parseInt(m[4]!, 10) - 1;
  return {
    r1: Math.min(r1, r2),
    r2: Math.max(r1, r2),
    c1: Math.min(c1, c2),
    c2: Math.max(c1, c2),
  };
}

function letterToCol(letter: string): number {
  let n = 0;
  for (const ch of letter) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

interface ChartCanvasProps {
  readonly kind: ChartKind;
  readonly width: number;
  readonly height: number;
  readonly categories: ReadonlyArray<string>;
  readonly series: ReadonlyArray<ChartSeries>;
}

function ChartCanvas(props: ChartCanvasProps): ReactNode {
  const { kind, width, height, categories, series } = props;
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
          horizontal={true}
        />
      );
    case "line":
      return <LineChart width={width} height={height} categories={categories} series={series} />;
    case "pie":
      return <PieChart width={width} height={height} series={series[0]!} categories={categories} />;
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
}

function ColumnOrBar(props: BaseChartProps & { horizontal: boolean }): ReactNode {
  const { width, height, categories, series, horizontal } = props;
  const padding = { top: 8, right: 12, bottom: 24, left: 36 };
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
            <line x1={xv(t)} x2={xv(t)} y1={padding.top} y2={padding.top + plotH} stroke="var(--divider)" />
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
              <rect
                key={`bar-${si}-${i}`}
                x={x}
                y={top}
                width={w}
                height={barH}
                fill={s.color}
                opacity={0.9}
              />
            );
          })
        )}
        <Legend series={series} x={width - 8} y={2} anchor="end" />
      </svg>
    );
  }

  return (
    <svg width={width} height={height}>
      {ticks.map((t) => (
        <g key={`gy-${t}`}>
          <line x1={padding.left} x2={padding.left + plotW} y1={ys(t)} y2={ys(t)} stroke="var(--divider)" />
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
            <rect
              key={`col-${si}-${i}`}
              x={left}
              y={top}
              width={barWidth}
              height={h}
              fill={s.color}
              opacity={0.9}
            />
          );
        })
      )}
      <Legend series={series} x={width - 8} y={2} anchor="end" />
    </svg>
  );
}

function LineChart(props: BaseChartProps): ReactNode {
  const { width, height, categories, series } = props;
  const padding = { top: 8, right: 12, bottom: 24, left: 36 };
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
          <line x1={padding.left} x2={padding.left + plotW} y1={ys(t)} y2={ys(t)} stroke="var(--divider)" />
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
            <path d={path} stroke={s.color} strokeWidth={1.6} fill="none" />
            {s.values.map((v, i) => (
              <circle key={`pt-${si}-${i}`} cx={xs(i)} cy={ys(v)} r={2} fill={s.color} />
            ))}
          </g>
        );
      })}
      <Legend series={series} x={width - 8} y={2} anchor="end" />
    </svg>
  );
}

function PieChart(props: {
  readonly width: number;
  readonly height: number;
  readonly series: ChartSeries;
  readonly categories: ReadonlyArray<string>;
}): ReactNode {
  const { width, height, series, categories } = props;
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
  let acc = -Math.PI / 2;
  return (
    <svg width={width} height={height}>
      {series.values.map((v, i) => {
        const value = Math.max(0, v);
        const angle = (value / total) * Math.PI * 2;
        const x0 = cx + radius * Math.cos(acc);
        const y0 = cy + radius * Math.sin(acc);
        acc += angle;
        const x1 = cx + radius * Math.cos(acc);
        const y1 = cy + radius * Math.sin(acc);
        const large = angle > Math.PI ? 1 : 0;
        const color = PALETTE[i % PALETTE.length]!;
        return (
          <path
            key={`slice-${i}`}
            d={`M ${cx} ${cy} L ${x0} ${y0} A ${radius} ${radius} 0 ${large} 1 ${x1} ${y1} Z`}
            fill={color}
            opacity={0.9}
          />
        );
      })}
      {categories.map((c, i) => {
        const color = PALETTE[i % PALETTE.length]!;
        return (
          <g key={`legend-${i}`} transform={`translate(${4}, ${10 + i * 12})`}>
            <rect width={8} height={8} fill={color} />
            <text x={12} y={7} fontSize={9} fill="var(--muted-foreground)">
              {truncate(c, 14)}
            </text>
          </g>
        );
      })}
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

function truncate(s: string, max: number): string {
  if (s.length <= max) return s;
  return `${s.slice(0, Math.max(1, max - 1))}…`;
}
