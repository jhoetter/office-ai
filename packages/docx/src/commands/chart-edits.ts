import { CommandError, type CommandHandler } from "@officeai/core";
import type {
  ChartPart,
  ChartSeries,
  ChartType,
  DocxDocument,
  DocxSnapshot,
} from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type {
  SetChartDataPayload,
  SetChartTitlePayload,
  SetChartTypePayload,
} from "./payloads.js";

/** Replace categories + series of an existing chart. */
export const setChartDataHandler: CommandHandler<SetChartDataPayload, DocxSnapshot> = {
  type: "docx:set-chart-data",
  apply(snapshot, payload, ctx) {
    const part = requireChart(snapshot, payload.chartPartPath);
    if (!Array.isArray(payload.categories) || payload.categories.length === 0) {
      throw new CommandError("invalid-payload", "categories must contain at least one entry");
    }
    if (!Array.isArray(payload.series) || payload.series.length === 0) {
      throw new CommandError("invalid-payload", "series must contain at least one entry");
    }
    for (let i = 0; i < payload.series.length; i++) {
      const s = payload.series[i]!;
      if (s.values.length !== payload.categories.length) {
        throw new CommandError(
          "invalid-payload",
          `series[${i}].values length ${s.values.length} ≠ categories length ${payload.categories.length}`
        );
      }
    }

    const nextPart: ChartPart = {
      ...part,
      categories: [...payload.categories],
      series: payload.series.map((s, i): ChartSeries => ({
        id: ctx.mintNodeId(),
        idx: i,
        ...(s.name !== undefined ? { name: s.name } : {}),
        values: [...s.values],
      })),
    };

    return finalize(snapshot, nextPart, "data");
  },
};

/** Set or clear the title of an existing chart. */
export const setChartTitleHandler: CommandHandler<SetChartTitlePayload, DocxSnapshot> = {
  type: "docx:set-chart-title",
  apply(snapshot, payload) {
    const part = requireChart(snapshot, payload.chartPartPath);
    const nextTitle = payload.title === null ? undefined : payload.title;
    if ((part.title ?? null) === (nextTitle ?? null)) {
      throw new CommandError("no-op", "title is unchanged");
    }
    const nextPart: ChartPart = nextTitle === undefined
      ? omitTitle(part)
      : { ...part, title: nextTitle };
    return finalize(snapshot, nextPart, "title");
  },
};

/** Switch the active plot type of an existing chart. */
export const setChartTypeHandler: CommandHandler<SetChartTypePayload, DocxSnapshot> = {
  type: "docx:set-chart-type",
  apply(snapshot, payload) {
    const part = requireChart(snapshot, payload.chartPartPath);
    if (part.chartType === payload.chartType) {
      throw new CommandError("no-op", `chart type already ${payload.chartType}`);
    }
    const nextPart: ChartPart = { ...part, chartType: payload.chartType as ChartType };
    return finalize(snapshot, nextPart, "type");
  },
};

function requireChart(snapshot: DocxSnapshot, chartPartPath: string): ChartPart {
  const part = snapshot.root.charts.get(chartPartPath);
  if (!part) {
    throw new CommandError("not-found", `no chart at ${chartPartPath}`);
  }
  return part;
}

function omitTitle(part: ChartPart): ChartPart {
  const { title: _drop, ...rest } = part;
  void _drop;
  return rest;
}

function finalize(
  snapshot: DocxSnapshot,
  nextPart: ChartPart,
  field: "data" | "title" | "type"
): { next: DocxSnapshot; diff: ReturnType<typeof buildDiff> } {
  const newCharts = new Map(snapshot.root.charts);
  newCharts.set(nextPart.partPath, nextPart);
  const nextDoc: DocxDocument = { ...snapshot.root, charts: newCharts };

  const next = evolveSnapshot(snapshot, nextDoc, {
    charts: withAddition(snapshot.dirty.charts, nextPart.partPath),
    contentTypes: true,
  });

  return {
    next,
    diff: buildDiff(snapshot.revision, next.revision, {
      kind: "node-updated",
      nodeId: nextPart.partPath,
      path: ["charts", nextPart.partPath],
      field,
      summary: `chart:${field}`,
    }),
  };
}

function withAddition(prev: ReadonlySet<string>, member: string): ReadonlySet<string> {
  const next = new Set(prev);
  next.add(member);
  return next;
}
