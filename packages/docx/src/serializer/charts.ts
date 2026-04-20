/**
 * DOCX chart serializer. Mirrors the PPTX chart-roundtrip pipeline:
 * for every chart part marked dirty we (re)build the typed
 * `<c:chartSpace>` XML via the shared `serializeChartXml` helper and
 * (re)materialise the embedded `Microsoft_Excel_WorksheetN.xlsx`
 * package that powers Office's "Edit Data" UI.
 *
 * The inline `<w:drawing>` envelope around each chart is serialized
 * by `serializeChartDrawing` in this module; the run-child serializer
 * dispatches to it when a `ChartDrawing` leaf has no captured `raw`.
 */

import { ooxml } from "@officeai/core";
import { buildChartGrid, buildEmbeddedXlsx } from "@officeai/xlsx";
import type { ChartDrawing, ChartPart, DocxSnapshot, Relationship } from "../model/types.js";
import { DocxSerializeError } from "./errors.js";

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/**
 * Re-emit every chart part in `dirty.charts`: rebuild the chart XML
 * from the typed model, register / refresh the embedded xlsx package,
 * and wire up content types + per-chart relationships so Office can
 * round-trip "Edit Data" against the embedded workbook.
 */
export async function serializeChartParts(
  container: ooxml.OoxmlContainer,
  snapshot: DocxSnapshot
): Promise<void> {
  if (snapshot.dirty.charts.size === 0) return;
  const contentTypes = ooxml.ContentTypes.load(container);
  let touchedContentTypes = false;
  for (const partPath of snapshot.dirty.charts) {
    const part = snapshot.root.charts.get(partPath);
    if (!part) continue;
    try {
      await serializeChartWithEmbedding(container, contentTypes, part);
      touchedContentTypes = true;
    } catch (err) {
      throw new DocxSerializeError("chart-failed", `Failed to serialize ${partPath}`, { cause: err });
    }
  }
  if (touchedContentTypes) contentTypes.writeBack(container);
}

async function serializeChartWithEmbedding(
  container: ooxml.OoxmlContainer,
  contentTypes: ooxml.ContentTypes,
  part: ChartPart
): Promise<void> {
  // Map the typed chart payload onto a 2D grid + the corresponding
  // A1-style cell references the chart XML will reach back into.
  const sheetName = part.embeddingSheetName ?? "Sheet1";
  const seriesForGrid = part.series.map((s) => ({
    values: [...s.values],
    ...(s.name !== undefined ? { name: s.name } : {}),
  }));
  const chartGrid = buildChartGrid([...part.categories], seriesForGrid, { sheetName });

  const embeddingPath = part.embeddingPartPath ?? mintEmbeddingPath(container, part.partPath);
  const built = await buildEmbeddedXlsx(chartGrid.grid, { sheetName });

  const chartRels = ooxml.RelationshipGraph.loadFor(container, part.partPath);
  const relTarget = ooxml.relativeTarget(part.partPath, embeddingPath);
  const added = ooxml.addEmbeddedPart({
    container,
    contentTypes,
    ownerRels: chartRels,
    partPath: embeddingPath,
    bytes: built.bytes,
    contentType: ooxml.CT_SPREADSHEETML_SHEET,
    relTarget,
    relType: ooxml.REL_TYPE_PACKAGE,
    ...(part.embeddingRelId ? { relId: part.embeddingRelId } : {}),
  });
  chartRels.writeBack(container);

  // Register the chart part itself in [Content_Types].xml.
  if (!contentTypes.hasOverride(absolutePartName(part.partPath))) {
    contentTypes.addOverride(absolutePartName(part.partPath), ooxml.CT_DRAWINGML_CHART);
  }

  const xml = ooxml.serializeChartXml(
    {
      chartType: part.chartType === "unsupported" ? "bar" : part.chartType,
      ...(part.title !== undefined ? { title: part.title } : {}),
      categories: part.categories,
      series: part.series.map((s) => ({
        idx: s.idx,
        ...(s.name !== undefined ? { name: s.name } : {}),
        values: s.values,
      })),
    },
    {
      embeddingRelId: added.relId,
      categoryRef: chartGrid.categoryRef,
      valueRefs: chartGrid.valueRefs,
      nameRefs: chartGrid.nameRefs,
    }
  );
  if (container.has(part.partPath)) container.writeText(part.partPath, xml);
  else container.addPart(part.partPath, new TextEncoder().encode(xml));
}

/**
 * Pick the next free `word/embeddings/Microsoft_Excel_WorksheetN.xlsx`
 * path. Mirrors Word's own naming convention so a saved-then-reopened
 * file stays human-readable inside the package.
 */
function mintEmbeddingPath(container: ooxml.OoxmlContainer, chartPartPath: string): string {
  const root = chartPartPath.startsWith("word/") ? "word" : (chartPartPath.split("/")[0] ?? "word");
  let n = 1;
  while (container.has(`${root}/embeddings/Microsoft_Excel_Worksheet${n}.xlsx`)) n++;
  return `${root}/embeddings/Microsoft_Excel_Worksheet${n}.xlsx`;
}

function absolutePartName(partPath: string): string {
  return partPath.startsWith("/") ? partPath : `/${partPath}`;
}

/**
 * Emit the inline `<w:drawing>` envelope for a freshly-authored chart
 * (one with no captured `raw`). The shape mirrors what Word writes for
 * `Insert > Chart`: a `<wp:inline>` whose `<a:graphicData>` URI points
 * at the chart schema and whose `<c:chart>` references the chart part
 * by relationship id.
 */
export function serializeChartDrawing(leaf: ChartDrawing): unknown {
  const inlineChildren: unknown[] = [];

  inlineChildren.push({
    "wp:extent": [],
    ":@": { "@_cx": String(leaf.cx), "@_cy": String(leaf.cy) },
  });
  inlineChildren.push({
    "wp:effectExtent": [],
    ":@": { "@_l": "0", "@_t": "0", "@_r": "0", "@_b": "0" },
  });

  const docPrAttrs: Record<string, string> = {
    "@_id": String(leaf.docPrId),
    "@_name": leaf.name,
  };
  if (leaf.descr !== undefined) docPrAttrs["@_descr"] = leaf.descr;
  inlineChildren.push({ "wp:docPr": [], ":@": docPrAttrs });

  const chartRef: Record<string, unknown> = {
    "c:chart": [],
    ":@": { "@_xmlns:c": C_NS, "@_xmlns:r": R_NS, "@_r:id": leaf.relId },
  };
  const graphicData: Record<string, unknown> = {
    "a:graphicData": [chartRef],
    ":@": { "@_uri": C_NS },
  };
  const graphic: Record<string, unknown> = {
    "a:graphic": [graphicData],
    ":@": { "@_xmlns:a": A_NS },
  };
  inlineChildren.push(graphic);

  const inlineAttrs: Record<string, string> = {
    "@_xmlns:wp": WP_NS,
    "@_distT": "0",
    "@_distB": "0",
    "@_distL": "0",
    "@_distR": "0",
  };
  return { "w:drawing": [{ "wp:inline": inlineChildren, ":@": inlineAttrs }] };
}

/**
 * Compute the chart relationship target as it should appear inside
 * `word/_rels/document.xml.rels` (relative to `word/`).
 */
export function chartRelTargetFor(chartPartPath: string): string {
  return chartPartPath.startsWith("word/") ? chartPartPath.slice("word/".length) : chartPartPath;
}

/** True iff `rels` already advertises a relationship pointing at `target`. */
export function findChartRel(rels: ReadonlyArray<Relationship>, target: string): Relationship | undefined {
  return rels.find(
    (r) =>
      r.type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" &&
      r.target === target
  );
}
