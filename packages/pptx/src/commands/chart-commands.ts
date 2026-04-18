import type { CommandHandler } from "@officeai/core";
import type {
  ChartPart,
  ChartSeries,
  ChartShape,
  ChartType,
  PptxPresentation,
  PptxSnapshot,
} from "../model/types.js";
import {
  buildDiff,
  evolveSnapshot,
  findShapeInSlide,
  findSlide,
  isChartShape,
  makeError,
} from "./helpers.js";
import type { SetChartDataPayload, SetChartTitlePayload, SetChartTypePayload } from "./payloads.js";

// ─── pptx:set-chart-title ────────────────────────────────────────────────

export const setChartTitleHandler: CommandHandler<SetChartTitlePayload, PptxSnapshot> = {
  type: "pptx:set-chart-title",
  apply(snapshot, payload) {
    const { chart, partPath } = resolveChart(snapshot, payload.slideIndex, payload.shapeId);
    const before = chart.title;
    const nextTitle = payload.title === null ? undefined : payload.title;
    if (before === nextTitle) {
      throw makeError("no-op", "chart title is unchanged");
    }
    const nextPart: ChartPart = nextTitle === undefined ? omitTitle(chart) : { ...chart, title: nextTitle };
    const root = withChart(snapshot.root, partPath, nextPart);
    const next = evolveSnapshot(snapshot, root, {
      slides: [snapshot.root.slides[payload.slideIndex].partPath],
      charts: [partPath],
    });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: payload.shapeId,
        path: ["slides", payload.slideIndex, "shapes", "*chart", "title"],
        field: "title",
        summary: nextTitle === undefined ? "(removed)" : `${before ?? "(none)"} → ${nextTitle}`,
      }),
    };
  },
};

// ─── pptx:set-chart-data ─────────────────────────────────────────────────

export const setChartDataHandler: CommandHandler<SetChartDataPayload, PptxSnapshot> = {
  type: "pptx:set-chart-data",
  apply(snapshot, payload, ctx) {
    if (payload.series.length === 0) {
      throw makeError("invalid-payload", "set-chart-data requires at least one series");
    }
    const expected = payload.categories.length;
    for (let i = 0; i < payload.series.length; i++) {
      const s = payload.series[i]!;
      if (s.values.length !== expected) {
        throw makeError(
          "invalid-payload",
          `series[${i}].values.length (${s.values.length}) must equal categories.length (${expected})`
        );
      }
    }
    const { chart, partPath } = resolveChart(snapshot, payload.slideIndex, payload.shapeId);
    const newSeries: ChartSeries[] = payload.series.map((s, i) => ({
      id: ctx.mintNodeId(),
      idx: i,
      ...(s.name !== undefined ? { name: s.name } : {}),
      values: [...s.values],
    }));
    const nextPart: ChartPart = {
      ...chart,
      categories: [...payload.categories],
      series: newSeries,
      // Drop opaque per-series caches: rebuilding from typed model.
      seriesRaw: new Map(),
    };
    const root = withChart(snapshot.root, partPath, nextPart);
    const next = evolveSnapshot(snapshot, root, {
      slides: [snapshot.root.slides[payload.slideIndex].partPath],
      charts: [partPath],
    });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: payload.shapeId,
        path: ["slides", payload.slideIndex, "shapes", "*chart", "data"],
        field: "data",
        summary: `${payload.categories.length} categories × ${payload.series.length} series`,
      }),
    };
  },
};

// ─── pptx:set-chart-type ─────────────────────────────────────────────────

const SUPPORTED_CHART_TYPES: ReadonlySet<ChartType> = new Set(["bar", "line", "pie", "area"]);

export const setChartTypeHandler: CommandHandler<SetChartTypePayload, PptxSnapshot> = {
  type: "pptx:set-chart-type",
  apply(snapshot, payload) {
    if (!SUPPORTED_CHART_TYPES.has(payload.chartType)) {
      throw makeError("invalid-payload", `chartType ${payload.chartType} is not supported`);
    }
    const { chart, partPath } = resolveChart(snapshot, payload.slideIndex, payload.shapeId);
    if (chart.chartType === payload.chartType) {
      throw makeError("no-op", `chart is already a ${payload.chartType} chart`);
    }
    const nextPart: ChartPart = { ...chart, chartType: payload.chartType };
    const root = withChart(snapshot.root, partPath, nextPart);
    const next = evolveSnapshot(snapshot, root, {
      slides: [snapshot.root.slides[payload.slideIndex].partPath],
      charts: [partPath],
    });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: payload.shapeId,
        path: ["slides", payload.slideIndex, "shapes", "*chart", "chartType"],
        field: "chartType",
        summary: `${chart.chartType} → ${payload.chartType}`,
      }),
    };
  },
};

// ─── helpers ─────────────────────────────────────────────────────────────

function resolveChart(
  snapshot: PptxSnapshot,
  slideIndex: number,
  shapeId: string
): { shape: ChartShape; chart: ChartPart; partPath: string } {
  const { slide } = findSlide(snapshot, slideIndex);
  const { shape } = findShapeInSlide(slide, shapeId);
  if (!isChartShape(shape)) {
    throw makeError("not-applicable", `shape ${shapeId} is not a chart`);
  }
  const part = snapshot.root.charts.get(shape.chartPartPath);
  if (!part) {
    throw makeError("unknown-target", `chart part not found: ${shape.chartPartPath}`);
  }
  return { shape, chart: part, partPath: shape.chartPartPath };
}

function withChart(root: PptxPresentation, partPath: string, next: ChartPart): PptxPresentation {
  const charts = new Map(root.charts);
  charts.set(partPath, next);
  return { ...root, charts };
}

function omitTitle(part: ChartPart): ChartPart {
  const { title: _drop, ...rest } = part;
  void _drop;
  return rest as ChartPart;
}
