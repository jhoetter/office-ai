import { CommandError, type CommandHandler } from "@officeai/core";
import type { ImageAnchor } from "../model/drawings.js";
import type { Sheet, SheetChart, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type {
  AddChartPayload,
  MoveChartPayload,
  RemoveChartPayload,
  ResizeChartPayload,
} from "./payloads.js";
import { parseRangeRef, resolveSheet } from "./validation.js";

const DEFAULT_WIDTH_PX = 480;
const DEFAULT_HEIGHT_PX = 280;

/**
 * `xlsx:add-chart` — drop a typed chart on a worksheet.
 *
 * The chart binds to a range of cells via `dataRange`; the editor
 * paints it from those cells at render time so changes to the
 * underlying values flow through automatically.
 *
 * Round-trip: the typed chart lives in the model; brand-new charts
 * do not (yet) re-emit DrawingML chart parts on save. Existing
 * `xl/charts/*.xml` parts round-trip verbatim via `opaqueParts`.
 */
export const addChartHandler: CommandHandler<AddChartPayload, XlsxSnapshot> = {
  type: "xlsx:add-chart",
  apply(snapshot, payload, ctx) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const range = parseRangeRef(payload.dataRange);
    const r1 = Math.min(range.start.row, range.end.row);
    const r2 = Math.max(range.start.row, range.end.row);
    const c1 = Math.min(range.start.col, range.end.col);
    const c2 = Math.max(range.start.col, range.end.col);
    if (r1 === r2 && c1 === c2) {
      throw new CommandError(
        "invalid-range",
        `add-chart range "${payload.dataRange}" must contain at least two cells`
      );
    }

    const anchor: ImageAnchor = payload.anchor ?? {
      fromRow: r1,
      fromCol: c2 + 1,
      fromOffsetXPx: 8,
      fromOffsetYPx: 0,
      widthPx: DEFAULT_WIDTH_PX,
      heightPx: DEFAULT_HEIGHT_PX,
      editAs: "oneCell",
    };

    const newChart: SheetChart = {
      id: ctx.mintNodeId(),
      kind: payload.kind,
      dataRange: payload.dataRange,
      hasHeaderRow: payload.hasHeaderRow !== false,
      hasCategoryColumn: payload.hasCategoryColumn !== false,
      ...(payload.title ? { title: payload.title } : {}),
      anchor,
    };

    const nextSheet: Sheet = { ...sheet, charts: [...sheet.charts, newChart] };
    const nextWorkbook: XlsxWorkbook = replaceSheet(snapshot.root, nextSheet);
    const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-inserted",
          nodeId: newChart.id,
          path: ["sheets", sheet.index, "charts", sheet.charts.length],
          summary: `Add ${payload.kind} chart to ${sheet.name}!${payload.dataRange}`,
          meta: { kind: payload.kind, range: payload.dataRange },
        },
      ]),
    };
  },
};

export const removeChartHandler: CommandHandler<RemoveChartPayload, XlsxSnapshot> = {
  type: "xlsx:remove-chart",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const idx = sheet.charts.findIndex((c) => c.id === payload.chartId);
    if (idx === -1) {
      throw new CommandError(
        "unknown-chart",
        `Sheet "${sheet.name}" has no chart with id ${payload.chartId}`
      );
    }
    const removed = sheet.charts[idx]!;
    const charts = sheet.charts.slice();
    charts.splice(idx, 1);
    const nextSheet: Sheet = { ...sheet, charts };
    const next = evolveSnapshot(snapshot, replaceSheet(snapshot.root, nextSheet), {
      sheets: [sheet.partPath],
    });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-deleted",
          nodeId: removed.id,
          path: ["sheets", sheet.index, "charts", idx],
          summary: `Remove ${removed.kind} chart from ${sheet.name}`,
        },
      ]),
    };
  },
};

export const moveChartHandler: CommandHandler<MoveChartPayload, XlsxSnapshot> = {
  type: "xlsx:move-chart",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const idx = sheet.charts.findIndex((c) => c.id === payload.chartId);
    if (idx === -1) {
      throw new CommandError(
        "unknown-chart",
        `Sheet "${sheet.name}" has no chart with id ${payload.chartId}`
      );
    }
    const chart = sheet.charts[idx]!;
    const nextAnchor: ImageAnchor = {
      ...chart.anchor,
      fromRow: Math.max(0, Math.floor(payload.fromRow)),
      fromCol: Math.max(0, Math.floor(payload.fromCol)),
      fromOffsetXPx: Math.max(0, payload.fromOffsetXPx),
      fromOffsetYPx: Math.max(0, payload.fromOffsetYPx),
    };
    const nextChart: SheetChart = { ...chart, anchor: nextAnchor };
    const charts = sheet.charts.slice();
    charts[idx] = nextChart;
    const next = evolveSnapshot(snapshot, replaceSheet(snapshot.root, { ...sheet, charts }), {
      sheets: [sheet.partPath],
    });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-updated",
          nodeId: chart.id,
          path: ["sheets", sheet.index, "charts", idx, "anchor"],
          field: "fromRow",
          summary: `Move chart on ${sheet.name}`,
        },
      ]),
    };
  },
};

export const resizeChartHandler: CommandHandler<ResizeChartPayload, XlsxSnapshot> = {
  type: "xlsx:resize-chart",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const idx = sheet.charts.findIndex((c) => c.id === payload.chartId);
    if (idx === -1) {
      throw new CommandError(
        "unknown-chart",
        `Sheet "${sheet.name}" has no chart with id ${payload.chartId}`
      );
    }
    const chart = sheet.charts[idx]!;
    const widthPx = Math.max(80, Math.round(payload.widthPx));
    const heightPx = Math.max(60, Math.round(payload.heightPx));
    const nextChart: SheetChart = {
      ...chart,
      anchor: { ...chart.anchor, widthPx, heightPx },
    };
    const charts = sheet.charts.slice();
    charts[idx] = nextChart;
    const next = evolveSnapshot(snapshot, replaceSheet(snapshot.root, { ...sheet, charts }), {
      sheets: [sheet.partPath],
    });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-updated",
          nodeId: chart.id,
          path: ["sheets", sheet.index, "charts", idx, "anchor"],
          field: "size",
          summary: `Resize chart on ${sheet.name}`,
        },
      ]),
    };
  },
};
