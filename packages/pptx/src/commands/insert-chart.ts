import type { CommandHandler } from "@officeai/core";
import type {
  ChartPart,
  ChartSeries,
  ChartShape,
  ChartType,
  ContentTypesSnap,
  OpaqueXml,
  PptxPresentation,
  PptxSnapshot,
  RelationshipsSnap,
  Slide,
} from "../model/types.js";
import { evolveSnapshot, findSlide, makeError, maxCNvPrId } from "./helpers.js";
import type { PptxInsertChartPayload } from "./payloads.js";

const CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const CHART_GRAPHIC_DATA_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const CT_DRAWINGML_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

const SUPPORTED_CHART_TYPES: ReadonlySet<ChartType> = new Set(["bar", "line", "pie", "area"]);

/**
 * Author a brand-new chart on a slide. Mirrors `docx:insert-chart`:
 *
 *   1. Mint `ppt/charts/chartN.xml` and a fresh slide rel (`chart`
 *      relationship) so `<p:graphicFrame>/<a:graphicData>/<c:chart>`
 *      resolves cleanly.
 *   2. Build a typed `ChartPart` (categories + series + chartType +
 *      optional title) plus a synthetic `chartSpaceRaw` containing an
 *      empty `<c:chart>` element. The serializer's chart pass uses
 *      this as the rebuild target, fills in the typed plotArea + title
 *      from the model, and authors the embedded `.xlsx` workbook so
 *      Office's "Edit Data" round-trips.
 *   3. Splice a typed `ChartShape` onto the slide.
 *   4. Mark dirty: slides (the new shape), charts (the new chart part),
 *      relationships (slide rels), contentTypes (chart override +
 *      embedded xlsx — registered by the serializer).
 */
export const insertChartHandler: CommandHandler<PptxInsertChartPayload, PptxSnapshot> = {
  type: "pptx:insert-chart",
  apply(snapshot, payload, ctx) {
    validatePayload(payload);

    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const slideRelsPath = relsPathFor(slide.partPath);

    const cx = Math.round(payload.cx ?? 4_572_000); // ~5 inches
    const cy = Math.round(payload.cy ?? 3_429_000); // ~3.75 inches

    const chartPartPath = mintChartPartPath(snapshot.root.charts);
    const slideRels = snapshot.relationships.get(slideRelsPath);
    const existingEntries = slideRels?.entries ?? [];
    const chartRelId = nextRelId(existingEntries.map((e) => e.id));
    const newSlideEntries = [
      ...existingEntries,
      {
        id: chartRelId,
        type: CHART_REL_TYPE,
        target: relativeFromRels(slideRelsPath, chartPartPath),
      },
    ];
    const newRelsMap = new Map(snapshot.relationships);
    newRelsMap.set(slideRelsPath, { relsPath: slideRelsPath, entries: newSlideEntries });

    const chartPart: ChartPart = {
      partPath: chartPartPath,
      contentType: CT_DRAWINGML_CHART,
      chartType: payload.chartType,
      ...(payload.title !== undefined ? { title: payload.title } : {}),
      categories: [...payload.categories],
      series: payload.series.map(
        (s, i): ChartSeries => ({
          id: ctx.mintNodeId(),
          idx: i,
          ...(s.name !== undefined ? { name: s.name } : {}),
          values: [...s.values],
        })
      ),
      chartSpaceRaw: emptyChartSpaceRaw(),
      plotAreaTailRaw: [],
      seriesRaw: new Map(),
      embeddingSheetName: payload.sheetName ?? "Sheet1",
    };

    const cNvPrId = maxCNvPrId(slide.shapes) + 1;
    const chartShape: ChartShape = {
      kind: "chart",
      id: ctx.mintNodeId(),
      cNvPrId,
      name: payload.name ?? `Chart ${cNvPrId}`,
      position: { xEmu: Math.round(payload.x), yEmu: Math.round(payload.y) },
      size: { cxEmu: cx, cyEmu: cy },
      chartRelId,
      chartPartPath,
      nvGraphicFramePrTail: [],
      graphicDataUri: CHART_GRAPHIC_DATA_URI,
    };

    const newSlide: Slide = { ...slide, shapes: [...slide.shapes, chartShape] };
    const newSlides = [...snapshot.root.slides];
    newSlides[sIdx] = newSlide;

    const newCharts = new Map(snapshot.root.charts);
    newCharts.set(chartPartPath, chartPart);

    const root: PptxPresentation = {
      ...snapshot.root,
      slides: newSlides,
      charts: newCharts,
    };

    const next = evolveSnapshot(
      snapshot,
      root,
      {
        slides: [slide.partPath],
        charts: [chartPartPath],
        relationships: [slideRelsPath],
        contentTypes: true,
      },
      {
        relationships: newRelsMap as ReadonlyMap<string, RelationshipsSnap>,
        contentTypes: registerChartOverride(snapshot.contentTypes, chartPartPath),
      }
    );

    return {
      next,
      diff: {
        format: "pptx",
        fromRevision: snapshot.revision,
        toRevision: next.revision,
        changes: [
          {
            kind: "node-inserted",
            nodeId: chartShape.id,
            path: ["slides", sIdx, "shapes", newSlide.shapes.length - 1] as ReadonlyArray<string | number>,
            summary: `+chart (${payload.chartType}, ${payload.categories.length} cats × ${payload.series.length} ser)`,
          },
          {
            kind: "part-added",
            path: [chartPartPath],
            summary: `+chart ${chartPartPath}`,
          },
        ],
      },
    };
  },
};

// ─── helpers ──────────────────────────────────────────────────────────────

function validatePayload(p: PptxInsertChartPayload): void {
  if (!SUPPORTED_CHART_TYPES.has(p.chartType)) {
    throw makeError("invalid-payload", `unsupported chartType: ${p.chartType}`);
  }
  if (!Array.isArray(p.categories) || p.categories.length === 0) {
    throw makeError("invalid-payload", "categories must contain at least one entry");
  }
  if (!Array.isArray(p.series) || p.series.length === 0) {
    throw makeError("invalid-payload", "series must contain at least one entry");
  }
  for (let i = 0; i < p.series.length; i++) {
    const s = p.series[i]!;
    if (!Array.isArray(s.values)) {
      throw makeError("invalid-payload", `series[${i}].values must be an array`);
    }
    if (s.values.length !== p.categories.length) {
      throw makeError(
        "invalid-payload",
        `series[${i}].values length ${s.values.length} ≠ categories length ${p.categories.length}`
      );
    }
  }
  if (p.cx !== undefined && (!Number.isFinite(p.cx) || p.cx <= 0)) {
    throw makeError("invalid-payload", `cx must be a positive number (got ${p.cx})`);
  }
  if (p.cy !== undefined && (!Number.isFinite(p.cy) || p.cy <= 0)) {
    throw makeError("invalid-payload", `cy must be a positive number (got ${p.cy})`);
  }
}

function mintChartPartPath(charts: ReadonlyMap<string, ChartPart>): string {
  let n = 1;
  while (charts.has(`ppt/charts/chart${n}.xml`)) n++;
  return `ppt/charts/chart${n}.xml`;
}

function nextRelId(existing: ReadonlyArray<string>): string {
  let max = 0;
  for (const id of existing) {
    const m = /^rId(\d+)$/.exec(id);
    if (m) {
      const n = Number(m[1]);
      if (n > max) max = n;
    }
  }
  return `rId${max + 1}`;
}

function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf("/");
  const dir = slash >= 0 ? partPath.slice(0, slash) : "";
  const file = slash >= 0 ? partPath.slice(slash + 1) : partPath;
  return `${dir}${dir ? "/" : ""}_rels/${file}.rels`;
}

function relativeFromRels(relsPath: string, targetAbsPath: string): string {
  const ownerDir = ownerDirOfRels(relsPath).split("/").filter(Boolean);
  const target = targetAbsPath.split("/").filter(Boolean);
  let i = 0;
  while (i < ownerDir.length && i < target.length - 1 && ownerDir[i] === target[i]) {
    i++;
  }
  const ups = ownerDir.length - i;
  const rest = target.slice(i);
  const parts: string[] = [];
  for (let k = 0; k < ups; k++) parts.push("..");
  for (const r of rest) parts.push(r);
  return parts.join("/");
}

function ownerDirOfRels(relsPath: string): string {
  const idx = relsPath.lastIndexOf("/_rels/");
  if (idx < 0) return "";
  return relsPath.slice(0, idx);
}

/**
 * A skeletal `<c:chartSpace>` node carrying just the empty `<c:chart>`
 * child the serializer's `serializeChartPartXml` walker needs as a
 * rebuild anchor. The plot area, title, series, and `<c:externalData>`
 * are filled in by the serializer from the typed model.
 */
function emptyChartSpaceRaw(): OpaqueXml {
  const emptyChart: OpaqueXml = {
    tag: "c:chart",
    attrs: {},
    rawAttrs: {},
    subtree: [],
  };
  return {
    tag: "c:chartSpace",
    attrs: {},
    rawAttrs: {},
    subtree: [{ "c:chart": emptyChart.subtree }],
  };
}

function registerChartOverride(contentTypes: ContentTypesSnap, chartPartPath: string): ContentTypesSnap {
  const partName = chartPartPath.startsWith("/") ? chartPartPath : `/${chartPartPath}`;
  if (contentTypes.overrides.some((o) => o.partName === partName)) return contentTypes;
  return {
    ...contentTypes,
    overrides: [...contentTypes.overrides, { partName, contentType: CT_DRAWINGML_CHART }],
  };
}
