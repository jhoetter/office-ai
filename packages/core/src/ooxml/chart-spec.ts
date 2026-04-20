/**
 * Typed projection of a `<c:chartSpace>` part shared between DOCX and
 * PPTX. The DrawingML chart schema is identical across both packages —
 * the only difference is how the consuming envelope (`<w:drawing>` vs
 * `<p:graphicFrame>`) references the chart by relationship. By
 * funnelling chart authoring through this module both packages emit
 * identically-structured XML, including the `<c:externalData>` link to
 * an embedded xlsx workbook that powers Office's "Edit Data" UI.
 *
 * This module knows nothing about the surrounding document — it takes
 * a fully resolved {@link ChartSpec} (categories, series, optional
 * title/type, optional embedded-data context) and returns a string of
 * `<c:chartSpace>` XML.
 */

import { serializeXml, type AttrMap } from "./xml.js";

const ATTR_KEY = ":@";

const CHART_NAMESPACES: Readonly<Record<string, string>> = {
  "xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
  "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
  "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
};

export type ChartType = "bar" | "line" | "pie" | "area";

export interface ChartSeriesSpec {
  /** Series ordering; emitted as both `c:idx` and `c:order`. */
  readonly idx: number;
  /** Optional series legend label. */
  readonly name?: string;
  /** Numeric values for each category. */
  readonly values: ReadonlyArray<number>;
}

export interface ChartSpec {
  readonly chartType: ChartType;
  readonly title?: string;
  readonly categories: ReadonlyArray<string>;
  readonly series: ReadonlyArray<ChartSeriesSpec>;
}

export interface ChartEmbeddingContext {
  /** `r:id` of the relationship from the chart part to its embedded workbook. */
  readonly embeddingRelId: string;
  /** A1-style reference for the categories range, e.g. `Sheet1!$A$2:$A$5`. */
  readonly categoryRef: string;
  /** A1-style references for each series' value range. */
  readonly valueRefs: ReadonlyArray<string>;
  /** A1-style references for each series' name cell. */
  readonly nameRefs: ReadonlyArray<string>;
}

/**
 * Build a `<c:chartSpace>` XML document from a typed {@link ChartSpec}.
 * When an embedding context is provided the cell-reference (`<c:f>`) and
 * `<c:externalData>` elements are written so Office can round-trip
 * "Edit Data" against the linked workbook.
 */
export function serializeChartXml(spec: ChartSpec, ctx?: ChartEmbeddingContext): string {
  const chartChildren: unknown[] = [];
  if (spec.title !== undefined) chartChildren.push(buildTitle(spec.title));
  chartChildren.push(buildAutoTitleDeleted(spec.title === undefined));
  chartChildren.push(buildPlotArea(spec, ctx));
  chartChildren.push(buildPlotVisOnly());

  const chartSpace: unknown[] = [makeEntry("c:chart", chartChildren)];
  if (ctx) chartSpace.push(buildExternalData(ctx.embeddingRelId));

  const root: Record<string, unknown> = { "c:chartSpace": chartSpace };
  root[ATTR_KEY] = makeRawAttrs(CHART_NAMESPACES);
  return serializeXml([root]);
}

// ── builders ──────────────────────────────────────────────────────────────

function buildTitle(title: string): Record<string, unknown> {
  const aT = makeEntry("a:t", [{ "#text": title }]);
  const aR = makeEntry("a:r", [aT]);
  const aP = makeEntry("a:p", [aR]);
  const cRich = makeEntry("c:rich", [aP]);
  const cTx = makeEntry("c:tx", [cRich]);
  const cOverlay = makeEntry("c:overlay", [], { val: "0" });
  return makeEntry("c:title", [cTx, cOverlay]);
}

function buildAutoTitleDeleted(deleted: boolean): Record<string, unknown> {
  return makeEntry("c:autoTitleDeleted", [], { val: deleted ? "1" : "0" });
}

function buildPlotVisOnly(): Record<string, unknown> {
  return makeEntry("c:plotVisOnly", [], { val: "1" });
}

function buildPlotArea(spec: ChartSpec, ctx?: ChartEmbeddingContext): Record<string, unknown> {
  const children: unknown[] = [makeEntry("c:layout", [])];
  children.push(buildChartTypeElement(spec, ctx));
  if (spec.chartType !== "pie") {
    children.push(buildAxis("c:catAx", 1, 2));
    children.push(buildAxis("c:valAx", 2, 1));
  }
  return makeEntry("c:plotArea", children);
}

function buildChartTypeElement(spec: ChartSpec, ctx?: ChartEmbeddingContext): Record<string, unknown> {
  const tag = chartTypeTag(spec.chartType);
  const children: unknown[] = [];
  if (spec.chartType === "bar") {
    children.push(makeEntry("c:barDir", [], { val: "col" }));
    children.push(makeEntry("c:grouping", [], { val: "clustered" }));
  } else if (spec.chartType === "line") {
    children.push(makeEntry("c:grouping", [], { val: "standard" }));
  } else if (spec.chartType === "area") {
    children.push(makeEntry("c:grouping", [], { val: "standard" }));
  }
  children.push(makeEntry("c:varyColors", [], { val: spec.chartType === "pie" ? "1" : "0" }));
  for (let i = 0; i < spec.series.length; i++) {
    children.push(buildSeries(spec, spec.series[i]!, ctx, i));
  }
  if (spec.chartType !== "pie") {
    children.push(makeEntry("c:axId", [], { val: "1" }));
    children.push(makeEntry("c:axId", [], { val: "2" }));
  }
  return makeEntry(tag, children);
}

function buildSeries(
  spec: ChartSpec,
  s: ChartSeriesSpec,
  ctx: ChartEmbeddingContext | undefined,
  seriesIndex: number
): Record<string, unknown> {
  const children: unknown[] = [];
  children.push(makeEntry("c:idx", [], { val: String(s.idx) }));
  children.push(makeEntry("c:order", [], { val: String(s.idx) }));
  if (s.name !== undefined) {
    const nameRef = ctx?.nameRefs[seriesIndex];
    children.push(buildSeriesName(s.name, nameRef));
  }
  if (spec.categories.length > 0) {
    children.push(buildCategoryRef(spec.categories, ctx?.categoryRef));
  }
  children.push(buildValueRef(s.values, ctx?.valueRefs[seriesIndex]));
  return makeEntry("c:ser", children);
}

function buildSeriesName(name: string, ref: string | undefined): Record<string, unknown> {
  if (!ref) {
    return makeEntry("c:tx", [makeEntry("c:v", [{ "#text": name }])]);
  }
  const cache = makeEntry("c:strCache", [
    makeEntry("c:ptCount", [], { val: "1" }),
    makeEntry("c:pt", [makeEntry("c:v", [{ "#text": name }])], { idx: "0" }),
  ]);
  return makeEntry("c:tx", [makeEntry("c:strRef", [makeEntry("c:f", [{ "#text": ref }]), cache])]);
}

function buildCategoryRef(
  categories: ReadonlyArray<string>,
  ref: string | undefined
): Record<string, unknown> {
  const ptCount = makeEntry("c:ptCount", [], { val: String(categories.length) });
  const pts: unknown[] = [ptCount];
  for (let i = 0; i < categories.length; i++) {
    pts.push(makeEntry("c:pt", [makeEntry("c:v", [{ "#text": categories[i] }])], { idx: String(i) }));
  }
  const formula = ref ?? `Sheet1!$A$2:$A$${categories.length + 1}`;
  const cache = makeEntry("c:strCache", pts);
  const refNode = makeEntry("c:strRef", [makeEntry("c:f", [{ "#text": formula }]), cache]);
  return makeEntry("c:cat", [refNode]);
}

function buildValueRef(values: ReadonlyArray<number>, ref: string | undefined): Record<string, unknown> {
  const ptCount = makeEntry("c:ptCount", [], { val: String(values.length) });
  const pts: unknown[] = [ptCount];
  for (let i = 0; i < values.length; i++) {
    pts.push(
      makeEntry("c:pt", [makeEntry("c:v", [{ "#text": String(values[i] ?? 0) }])], { idx: String(i) })
    );
  }
  const formula = ref ?? `Sheet1!$B$2:$B$${values.length + 1}`;
  const cache = makeEntry("c:numCache", [makeEntry("c:formatCode", [{ "#text": "General" }]), ...pts]);
  const refNode = makeEntry("c:numRef", [makeEntry("c:f", [{ "#text": formula }]), cache]);
  return makeEntry("c:val", [refNode]);
}

function buildAxis(tag: "c:catAx" | "c:valAx", axId: number, crossAx: number): Record<string, unknown> {
  return makeEntry(tag, [
    makeEntry("c:axId", [], { val: String(axId) }),
    makeEntry("c:scaling", [makeEntry("c:orientation", [], { val: "minMax" })]),
    makeEntry("c:delete", [], { val: "0" }),
    makeEntry("c:axPos", [], { val: tag === "c:catAx" ? "b" : "l" }),
    makeEntry("c:crossAx", [], { val: String(crossAx) }),
  ]);
}

function buildExternalData(relId: string): Record<string, unknown> {
  const auto = makeEntry("c:autoUpdate", [], { val: "0" });
  const entry: Record<string, unknown> = { "c:externalData": [auto] };
  entry[ATTR_KEY] = makeRawAttrs({ "r:id": relId });
  return entry;
}

function chartTypeTag(t: ChartType): string {
  switch (t) {
    case "bar":
      return "c:barChart";
    case "line":
      return "c:lineChart";
    case "pie":
      return "c:pieChart";
    case "area":
      return "c:areaChart";
  }
}

// ── XML entry helpers (mirrors PPTX serializer's local helpers) ──────────

function makeEntry(
  tag: string,
  children: ReadonlyArray<unknown>,
  attrs?: Record<string, string>
): Record<string, unknown> {
  const e: Record<string, unknown> = { [tag]: children };
  if (attrs && Object.keys(attrs).length > 0) {
    e[ATTR_KEY] = makeRawAttrs(attrs);
  }
  return e;
}

function makeRawAttrs(attrs: Record<string, string>): AttrMap {
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(attrs)) {
    out[`@_${k}`] = v;
  }
  return out;
}
