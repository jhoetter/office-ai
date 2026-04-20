import { CommandError, ooxml, type CommandHandler } from "@officeai/core";
import type { ImageAnchor } from "../model/drawings.js";
import type { Sheet, SheetChart, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet, type PartialDirtyFlags } from "./helpers.js";
import type {
  AddChartPayload,
  MoveChartPayload,
  RemoveChartPayload,
  ResizeChartPayload,
  UpdateChartPayload,
} from "./payloads.js";
import { parseRangeRef, resolveSheet } from "./validation.js";

const DEFAULT_WIDTH_PX = 480;
const DEFAULT_HEIGHT_PX = 280;

/**
 * Chart mutations always touch the same triplet: the sheet itself
 * (so the worksheet XML can splice in the `<drawing>` ref), the
 * sheet's drawing part (where the `xdr:graphicFrame` lives), the
 * sheet's rels (for the drawing relationship) and content-types
 * (for the chart override entries the serializer materialises from
 * the container after the drawing pass). We bundle them up here so
 * the three command handlers stay in lock-step.
 */
function chartDirtyPatch(sheet: Sheet): PartialDirtyFlags {
  return {
    sheets: [sheet.partPath],
    drawings: [sheet.partPath],
    sheetRels: [ooxml.RelationshipGraph.relsPathFor(sheet.partPath)],
    contentTypes: true,
  };
}

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
      ...(payload.palette !== undefined ? { palette: payload.palette } : {}),
      ...(payload.showLegend !== undefined ? { showLegend: payload.showLegend } : {}),
      ...(payload.showDataLabels !== undefined ? { showDataLabels: payload.showDataLabels } : {}),
      ...(payload.showGridlines !== undefined ? { showGridlines: payload.showGridlines } : {}),
      ...(payload.xAxisTitle ? { xAxisTitle: payload.xAxisTitle } : {}),
      ...(payload.yAxisTitle ? { yAxisTitle: payload.yAxisTitle } : {}),
    };

    const nextSheet: Sheet = { ...sheet, charts: [...sheet.charts, newChart] };
    const nextWorkbook: XlsxWorkbook = replaceSheet(snapshot.root, nextSheet);
    const next = evolveSnapshot(snapshot, nextWorkbook, chartDirtyPatch(sheet));
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
    const next = evolveSnapshot(snapshot, replaceSheet(snapshot.root, nextSheet), chartDirtyPatch(sheet));
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
    const next = evolveSnapshot(snapshot, replaceSheet(snapshot.root, { ...sheet, charts }), chartDirtyPatch(sheet));
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

/**
 * `xlsx:update-chart` — patch a chart's typed properties in place.
 *
 * Mirrors {@link moveChartHandler}'s lookup-then-replace shape, but
 * touches `kind` / `dataRange` / `title` / `hasHeaderRow` /
 * `hasCategoryColumn` instead of the anchor. Only fields explicitly
 * present in the payload are written; the handler rejects the
 * command when nothing would change so palette typos surface as a
 * visible error rather than a silent no-op.
 */
export const updateChartHandler: CommandHandler<UpdateChartPayload, XlsxSnapshot> = {
  type: "xlsx:update-chart",
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

    const hasKind = payload.kind !== undefined;
    const hasRange = payload.dataRange !== undefined;
    const hasTitle = payload.title !== undefined;
    const hasHeader = payload.hasHeaderRow !== undefined;
    const hasCategory = payload.hasCategoryColumn !== undefined;
    const hasPalette = payload.palette !== undefined;
    const hasLegend = payload.showLegend !== undefined;
    const hasDataLabels = payload.showDataLabels !== undefined;
    const hasGridlines = payload.showGridlines !== undefined;
    const hasXAxisTitle = payload.xAxisTitle !== undefined;
    const hasYAxisTitle = payload.yAxisTitle !== undefined;
    if (
      !hasKind &&
      !hasRange &&
      !hasTitle &&
      !hasHeader &&
      !hasCategory &&
      !hasPalette &&
      !hasLegend &&
      !hasDataLabels &&
      !hasGridlines &&
      !hasXAxisTitle &&
      !hasYAxisTitle
    ) {
      throw new CommandError(
        "no-op",
        `update-chart for ${payload.chartId} did not specify any field to change`
      );
    }

    if (hasRange) {
      const range = parseRangeRef(payload.dataRange!);
      const r1 = Math.min(range.start.row, range.end.row);
      const r2 = Math.max(range.start.row, range.end.row);
      const c1 = Math.min(range.start.col, range.end.col);
      const c2 = Math.max(range.start.col, range.end.col);
      if (r1 === r2 && c1 === c2) {
        throw new CommandError(
          "invalid-range",
          `update-chart range "${payload.dataRange}" must contain at least two cells`
        );
      }
    }

    const baseTitle: string | undefined = hasTitle
      ? payload.title === null || payload.title === ""
        ? undefined
        : payload.title!
      : chart.title;
    const nextXAxisTitle: string | undefined = hasXAxisTitle
      ? payload.xAxisTitle === null || payload.xAxisTitle === ""
        ? undefined
        : payload.xAxisTitle!
      : chart.xAxisTitle;
    const nextYAxisTitle: string | undefined = hasYAxisTitle
      ? payload.yAxisTitle === null || payload.yAxisTitle === ""
        ? undefined
        : payload.yAxisTitle!
      : chart.yAxisTitle;
    const nextPalette = hasPalette
      ? payload.palette === null
        ? undefined
        : payload.palette!
      : chart.palette;
    const nextShowLegend = hasLegend
      ? payload.showLegend === null
        ? undefined
        : payload.showLegend!
      : chart.showLegend;
    const nextShowDataLabels = hasDataLabels
      ? payload.showDataLabels === null
        ? undefined
        : payload.showDataLabels!
      : chart.showDataLabels;
    const nextShowGridlines = hasGridlines
      ? payload.showGridlines === null
        ? undefined
        : payload.showGridlines!
      : chart.showGridlines;
    const nextChart: SheetChart = {
      ...chart,
      ...(hasKind ? { kind: payload.kind! } : {}),
      ...(hasRange ? { dataRange: payload.dataRange! } : {}),
      ...(hasHeader ? { hasHeaderRow: payload.hasHeaderRow! } : {}),
      ...(hasCategory ? { hasCategoryColumn: payload.hasCategoryColumn! } : {}),
      ...(baseTitle !== undefined ? { title: baseTitle } : {}),
      ...(nextPalette !== undefined ? { palette: nextPalette } : {}),
      ...(nextShowLegend !== undefined ? { showLegend: nextShowLegend } : {}),
      ...(nextShowDataLabels !== undefined ? { showDataLabels: nextShowDataLabels } : {}),
      ...(nextShowGridlines !== undefined ? { showGridlines: nextShowGridlines } : {}),
      ...(nextXAxisTitle !== undefined ? { xAxisTitle: nextXAxisTitle } : {}),
      ...(nextYAxisTitle !== undefined ? { yAxisTitle: nextYAxisTitle } : {}),
    };
    if (baseTitle === undefined) delete (nextChart as { title?: string }).title;
    if (nextPalette === undefined) delete (nextChart as { palette?: string }).palette;
    if (nextShowLegend === undefined) delete (nextChart as { showLegend?: boolean }).showLegend;
    if (nextShowDataLabels === undefined)
      delete (nextChart as { showDataLabels?: boolean }).showDataLabels;
    if (nextShowGridlines === undefined)
      delete (nextChart as { showGridlines?: boolean }).showGridlines;
    if (nextXAxisTitle === undefined) delete (nextChart as { xAxisTitle?: string }).xAxisTitle;
    if (nextYAxisTitle === undefined) delete (nextChart as { yAxisTitle?: string }).yAxisTitle;

    if (chartShallowEqual(chart, nextChart)) {
      throw new CommandError(
        "no-op",
        `update-chart for ${payload.chartId} did not actually change anything`
      );
    }

    const charts = sheet.charts.slice();
    charts[idx] = nextChart;
    const next = evolveSnapshot(
      snapshot,
      replaceSheet(snapshot.root, { ...sheet, charts }),
      chartDirtyPatch(sheet)
    );
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-updated",
          nodeId: chart.id,
          path: ["sheets", sheet.index, "charts", idx],
          field: "properties",
          summary: `Update chart on ${sheet.name}`,
        },
      ]),
    };
  },
};

/**
 * Detect a no-op `update-chart` after the patch is applied. Compares
 * the typed-property surface (`kind`, `dataRange`, `title`,
 * `hasHeaderRow`, `hasCategoryColumn`) — anchors are out of scope
 * because they're owned by `move-chart` / `resize-chart`.
 */
function chartShallowEqual(a: SheetChart, b: SheetChart): boolean {
  return (
    a.kind === b.kind &&
    a.dataRange === b.dataRange &&
    a.title === b.title &&
    a.hasHeaderRow === b.hasHeaderRow &&
    a.hasCategoryColumn === b.hasCategoryColumn &&
    a.palette === b.palette &&
    a.showLegend === b.showLegend &&
    a.showDataLabels === b.showDataLabels &&
    a.showGridlines === b.showGridlines &&
    a.xAxisTitle === b.xAxisTitle &&
    a.yAxisTitle === b.yAxisTitle
  );
}

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
    const next = evolveSnapshot(snapshot, replaceSheet(snapshot.root, { ...sheet, charts }), chartDirtyPatch(sheet));
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
