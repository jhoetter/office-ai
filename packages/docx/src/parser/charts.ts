/**
 * Chart parser for DOCX. Mirrors the PPTX `parseChartPart` shape so the
 * shared chart-spec helpers in `@officeai/core` can be applied later
 * during serialization. We extract the typed projection (categories,
 * series, type, title) and let everything else round-trip via the
 * container's part cache when the chart isn't dirtied.
 */

import { ooxml, type IdMinter } from "@officeai/core";
import type { ChartPart, ChartSeries, ChartType, Relationship } from "../model/types.js";
import { DocxParseError } from "./errors.js";
import { attrOf, elementEntries, findElementEntry } from "./xml-helpers.js";

const REL_TYPE_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const REL_TYPE_PACKAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
const REL_TYPE_OLE_OBJECT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";

const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

const CHART_TYPE_TAGS: ReadonlyArray<{ tag: string; type: ChartType }> = [
  { tag: "c:barChart", type: "bar" },
  { tag: "c:bar3DChart", type: "bar" },
  { tag: "c:lineChart", type: "line" },
  { tag: "c:line3DChart", type: "line" },
  { tag: "c:pieChart", type: "pie" },
  { tag: "c:pie3DChart", type: "pie" },
  { tag: "c:doughnutChart", type: "pie" },
  { tag: "c:areaChart", type: "area" },
  { tag: "c:area3DChart", type: "area" },
];

/**
 * Discover every `word/charts/chart*.xml` part referenced from
 * `word/document.xml.rels`, parse each into a typed {@link ChartPart},
 * and return the path → part map for {@link DocxDocument.charts}.
 *
 * Charts referenced only from header/footer parts are not currently
 * lifted into the typed map (they round-trip verbatim via the
 * container's part cache, which preserves byte-equality for
 * untouched documents).
 */
export function parseChartParts(
  container: ooxml.OoxmlContainer,
  relationships: ReadonlyMap<string, ReadonlyArray<Relationship>>,
  mintNodeId: IdMinter
): ReadonlyMap<string, ChartPart> {
  const out = new Map<string, ChartPart>();
  const docRels = relationships.get("word/document.xml") ?? [];
  for (const rel of docRels) {
    if (rel.type !== REL_TYPE_CHART) continue;
    const partPath = resolveTarget("word/document.xml", rel.target);
    if (!container.has(partPath)) continue;
    if (out.has(partPath)) continue;
    const xml = container.readText(partPath);
    const embedding = readChartEmbedding(container, partPath);
    out.set(partPath, parseChartPart(partPath, xml, mintNodeId, embedding));
  }
  return out;
}

interface ChartEmbedding {
  embeddingPartPath?: string;
  embeddingRelId?: string;
}

function readChartEmbedding(container: ooxml.OoxmlContainer, chartPartPath: string): ChartEmbedding {
  const relsPath = ooxml.RelationshipGraph.relsPathFor(chartPartPath);
  if (!container.has(relsPath)) return {};
  const graph = ooxml.RelationshipGraph.loadFor(container, chartPartPath);
  const pkg = graph.relationships.find((r) => r.type === REL_TYPE_PACKAGE || r.type === REL_TYPE_OLE_OBJECT);
  if (!pkg) return {};
  return { embeddingPartPath: resolveTarget(chartPartPath, pkg.target), embeddingRelId: pkg.id };
}

function parseChartPart(
  partPath: string,
  xml: string,
  mintNodeId: IdMinter,
  embedding: ChartEmbedding
): ChartPart {
  let tree: unknown[];
  try {
    tree = ooxml.parseXml(xml) as unknown[];
  } catch (err) {
    throw new DocxParseError("invalid-xml", `Failed to parse ${partPath}`, { partPath, cause: err });
  }
  const chartSpace = findElementEntry(tree, "c:chartSpace");
  if (!chartSpace) {
    throw new DocxParseError("invalid-xml", `Missing <c:chartSpace> in ${partPath}`, { partPath });
  }
  const chartSpaceChildren = (chartSpace["c:chartSpace"] as unknown[] | undefined) ?? [];
  const chart = findElementEntry(chartSpaceChildren, "c:chart");
  const chartChildren = chart ? ((chart["c:chart"] as unknown[] | undefined) ?? []) : [];
  const title = readChartTitle(chartChildren);
  const plotArea = findElementEntry(chartChildren, "c:plotArea");
  const plotAreaChildren = plotArea ? ((plotArea["c:plotArea"] as unknown[] | undefined) ?? []) : [];

  let chartType: ChartType = "unsupported";
  let chartTypeEntry: Record<string, unknown> | null = null;
  for (const cand of CHART_TYPE_TAGS) {
    const e = findElementEntry(plotAreaChildren, cand.tag);
    if (e) {
      chartType = cand.type;
      chartTypeEntry = e;
      break;
    }
  }

  const series: ChartSeries[] = [];
  let categories: ReadonlyArray<string> = [];
  if (chartTypeEntry) {
    const ctChildren = (chartTypeEntry[ooxml.getTag(chartTypeEntry)] as unknown[] | undefined) ?? [];
    for (const sEntry of elementEntries(ctChildren)) {
      if (ooxml.getTag(sEntry) !== "c:ser") continue;
      const sChildren = (sEntry["c:ser"] as unknown[] | undefined) ?? [];
      const idxEntry = findElementEntry(sChildren, "c:idx");
      const idx = idxEntry ? Number(attrOf(idxEntry, "val") ?? "0") : series.length;
      const txEntry = findElementEntry(sChildren, "c:tx");
      const name = txEntry ? readSeriesText(txEntry) : undefined;
      const valEntry = findElementEntry(sChildren, "c:val");
      const values = valEntry ? readNumericCache(valEntry, "c:val") : [];
      const catEntry = findElementEntry(sChildren, "c:cat");
      if (categories.length === 0 && catEntry) {
        categories = readStringCache(catEntry, "c:cat");
      }
      series.push({
        id: mintNodeId(),
        idx,
        ...(name !== undefined ? { name } : {}),
        values,
      });
    }
  }

  return {
    partPath,
    contentType: CHART_CONTENT_TYPE,
    chartType,
    ...(title !== undefined ? { title } : {}),
    categories,
    series,
    ...(embedding.embeddingPartPath ? { embeddingPartPath: embedding.embeddingPartPath } : {}),
    ...(embedding.embeddingRelId ? { embeddingRelId: embedding.embeddingRelId } : {}),
  };
}

function readChartTitle(chartChildren: ReadonlyArray<unknown>): string | undefined {
  const titleEntry = findElementEntry(chartChildren, "c:title");
  if (!titleEntry) return undefined;
  const tx = findElementEntry((titleEntry["c:title"] as unknown[] | undefined) ?? [], "c:tx");
  if (!tx) return undefined;
  const rich = findElementEntry((tx["c:tx"] as unknown[] | undefined) ?? [], "c:rich");
  const root = rich ?? findElementEntry((tx["c:tx"] as unknown[] | undefined) ?? [], "c:strRef");
  if (!root) return undefined;
  return collectTitleText(root).trim() || undefined;
}

function collectTitleText(node: Record<string, unknown>): string {
  const tag = ooxml.getTag(node);
  const children = (node[tag] as unknown[] | undefined) ?? [];
  let text = "";
  for (const c of children) {
    if (!c || typeof c !== "object") continue;
    const e = c as Record<string, unknown>;
    if ("#text" in e && typeof e["#text"] === "string") {
      text += e["#text"];
      continue;
    }
    text += collectTitleText(e);
  }
  return text;
}

function readSeriesText(txEntry: Record<string, unknown>): string | undefined {
  const children = (txEntry["c:tx"] as unknown[] | undefined) ?? [];
  const v = findElementEntry(children, "c:v");
  if (v) return readVText(v);
  const strRef = findElementEntry(children, "c:strRef");
  if (strRef) {
    const cache = findElementEntry((strRef["c:strRef"] as unknown[] | undefined) ?? [], "c:strCache");
    if (cache) {
      const pts = (cache["c:strCache"] as unknown[] | undefined) ?? [];
      const first = findElementEntry(pts, "c:pt");
      if (first) {
        const fv = findElementEntry((first["c:pt"] as unknown[] | undefined) ?? [], "c:v");
        if (fv) return readVText(fv);
      }
    }
  }
  return undefined;
}

function readVText(v: Record<string, unknown>): string {
  const children = (v["c:v"] as unknown[] | undefined) ?? [];
  for (const c of children) {
    if (c && typeof c === "object" && "#text" in (c as Record<string, unknown>)) {
      const t = (c as Record<string, unknown>)["#text"];
      if (typeof t === "string") return t;
    }
  }
  return "";
}

function readNumericCache(entry: Record<string, unknown>, tag: string): number[] {
  const children = (entry[tag] as unknown[] | undefined) ?? [];
  const numRef = findElementEntry(children, "c:numRef");
  const cacheRoot = numRef
    ? findElementEntry((numRef["c:numRef"] as unknown[] | undefined) ?? [], "c:numCache")
    : findElementEntry(children, "c:numCache");
  if (!cacheRoot) return [];
  const cacheChildren = (cacheRoot["c:numCache"] as unknown[] | undefined) ?? [];
  const out: number[] = [];
  for (const pt of elementEntries(cacheChildren)) {
    if (ooxml.getTag(pt) !== "c:pt") continue;
    const idxAttr = attrOf(pt, "idx");
    const idx = idxAttr !== undefined ? Number(idxAttr) : out.length;
    const vEntry = findElementEntry((pt["c:pt"] as unknown[] | undefined) ?? [], "c:v");
    if (!vEntry) continue;
    const txt = readVText(vEntry);
    const num = Number(txt);
    if (!Number.isFinite(num)) continue;
    while (out.length <= idx) out.push(0);
    out[idx] = num;
  }
  return out;
}

function readStringCache(entry: Record<string, unknown>, tag: string): string[] {
  const children = (entry[tag] as unknown[] | undefined) ?? [];
  const strRef = findElementEntry(children, "c:strRef");
  const cacheRoot = strRef
    ? findElementEntry((strRef["c:strRef"] as unknown[] | undefined) ?? [], "c:strCache")
    : findElementEntry(children, "c:strCache");
  if (!cacheRoot) return [];
  const cacheChildren = (cacheRoot["c:strCache"] as unknown[] | undefined) ?? [];
  const out: string[] = [];
  for (const pt of elementEntries(cacheChildren)) {
    if (ooxml.getTag(pt) !== "c:pt") continue;
    const idxAttr = attrOf(pt, "idx");
    const idx = idxAttr !== undefined ? Number(idxAttr) : out.length;
    const vEntry = findElementEntry((pt["c:pt"] as unknown[] | undefined) ?? [], "c:v");
    if (!vEntry) continue;
    const txt = readVText(vEntry);
    while (out.length <= idx) out.push("");
    out[idx] = txt;
  }
  return out;
}

/**
 * Resolve a relationship `target` (which may be relative or absolute)
 * against the owning part path, returning a normalized OOXML part
 * path with no leading slash.
 */
function resolveTarget(ownerPartPath: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const ownerDir = ownerPartPath.includes("/")
    ? ownerPartPath.slice(0, ownerPartPath.lastIndexOf("/") + 1)
    : "";
  const segments = (ownerDir + target).split("/");
  const resolved: string[] = [];
  for (const seg of segments) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") {
      resolved.pop();
      continue;
    }
    resolved.push(seg);
  }
  return resolved.join("/");
}
