import { ooxml } from "@officeai/core";
import type { ChartKind, SheetChart } from "../model/types.js";

/**
 * Content-type for the DrawingML chart part. Excel and LibreOffice
 * both require this `<Override>` entry — without it the chart parts
 * are loaded as opaque XML and the chart goes invisible.
 */
export const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

/**
 * Relationship type for `xl/drawings/drawingN.xml` → `xl/charts/chartN.xml`.
 * Same `relationships/chart` namespace as Office uses everywhere.
 */
export const CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

const DRAWINGML_CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const RELS_OFFICE_DOC_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/**
 * Serialise a typed {@link SheetChart} into a `<c:chartSpace>` payload
 * suitable for `xl/charts/chartN.xml`.
 *
 * Why we hand-write the DrawingML rather than building it from a tree:
 * the chart part is small (≤ 2 KB for the kinds we author) and almost
 * entirely shaped by the choice of `kind`. A purpose-built emitter is
 * easier to reason about than a generic `XmlElement` builder when the
 * point is to match Excel's exact element ordering — the Office app
 * is famously picky about element order inside `<c:plotArea>` (see
 * MS-OE376 §21.2). We keep the emitter close to that canonical order
 * so files round-trip into Excel and LibreOffice without warnings.
 *
 * Series semantics:
 *  - `dataRange` is parsed by the *caller* into a `SheetName!A1:Cn`
 *    string we re-quote here.
 *  - When `hasHeaderRow` is `true`, the first row contributes series
 *    titles; the remaining rows are data.
 *  - When `hasCategoryColumn` is `true`, the first column contributes
 *    category labels; the remaining columns are series.
 *
 * The function does NOT bake values into the chart part (no `<c:numLit>` /
 * `<c:strLit>` caches) — Excel will recompute them from the source
 * cells on open. Skipping the value cache keeps the chart payload tiny
 * and keeps the round-trip honest: the only place a chart's data
 * lives is the worksheet cells the chart points at.
 */
export function serializeChartPart(chart: SheetChart, ownerSheetName?: string): string {
  // If the model stored a sheet-unqualified `A1:C5` range we
  // synthesise the qualifier from the chart's owning sheet. Authored
  // charts go through the unqualified path because `xlsx:add-chart`
  // takes a `dataRange` relative to its `sheet` payload field.
  const ref = chart.dataRange.includes("!")
    ? chart.dataRange
    : ownerSheetName
      ? `${quoteSheetName(ownerSheetName)}!${chart.dataRange}`
      : chart.dataRange;
  const range = parseQualifiedRange(ref);
  if (!range) {
    // We still emit a syntactically valid chart shell so the file
    // opens. The chart will just render empty inside Office.
    return wrapChartSpace(emptyChart(chart.kind));
  }

  const { sheet, startRow, startCol, endRow, endCol } = range;
  const dataStartRow = chart.hasHeaderRow ? startRow + 1 : startRow;
  const dataStartCol = chart.hasCategoryColumn ? startCol + 1 : startCol;

  // Series = one per data column (or one row when the data is a
  // single column with `hasCategoryColumn = false`). We pick the
  // "more cells means more series" axis so a 5×3 range with header +
  // category produces 2 series of 4 points, not 4 series of 2 points.
  const seriesCount = Math.max(1, endCol - dataStartCol + 1);
  const series: SeriesSpec[] = [];
  for (let s = 0; s < seriesCount; s++) {
    const col = dataStartCol + s;
    const valuesRange = `${sheet}!${a1(dataStartRow, col)}:${a1(endRow, col)}`;
    const titleRef = chart.hasHeaderRow ? `${sheet}!${a1(startRow, col)}` : undefined;
    const categoriesRef = chart.hasCategoryColumn
      ? `${sheet}!${a1(dataStartRow, startCol)}:${a1(endRow, startCol)}`
      : undefined;
    series.push({
      idx: s,
      valuesRange,
      ...(titleRef ? { titleRef } : {}),
      ...(categoriesRef ? { categoriesRef } : {}),
    });
  }

  const body = buildChartBody(chart, series);
  return wrapChartSpace(body);
}

interface SeriesSpec {
  readonly idx: number;
  readonly valuesRange: string;
  readonly titleRef?: string;
  readonly categoriesRef?: string;
}

function buildChartBody(chart: SheetChart, series: ReadonlyArray<SeriesSpec>): string {
  const titleXml = chart.title
    ? `<c:title>` +
      `<c:tx><c:rich>` +
      `<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" wrap="square" anchor="ctr" anchorCtr="1"/>` +
      `<a:lstStyle/>` +
      `<a:p><a:r><a:rPr lang="en-US" sz="1400" b="0"/>` +
      `<a:t>${escapeXml(chart.title)}</a:t></a:r></a:p>` +
      `</c:rich></c:tx>` +
      `<c:overlay val="0"/>` +
      `</c:title>`
    : "";

  const plotArea = buildPlotArea(chart.kind, series);

  return (
    `<c:chart>` +
    titleXml +
    `<c:autoTitleDeleted val="${chart.title ? 0 : 1}"/>` +
    plotArea +
    `<c:legend>` +
    `<c:legendPos val="r"/>` +
    `<c:overlay val="0"/>` +
    `</c:legend>` +
    `<c:plotVisOnly val="1"/>` +
    `<c:dispBlanksAs val="gap"/>` +
    `</c:chart>`
  );
}

function buildPlotArea(kind: ChartKind, series: ReadonlyArray<SeriesSpec>): string {
  const seriesXml = series.map((s) => buildSeries(kind, s)).join("");
  switch (kind) {
    case "column":
      return (
        `<c:plotArea>` +
        `<c:layout/>` +
        `<c:barChart>` +
        `<c:barDir val="col"/>` +
        `<c:grouping val="clustered"/>` +
        `<c:varyColors val="0"/>` +
        seriesXml +
        `<c:axId val="111111111"/>` +
        `<c:axId val="222222222"/>` +
        `</c:barChart>` +
        defaultCatAxis("111111111", "222222222") +
        defaultValAxis("222222222", "111111111") +
        `</c:plotArea>`
      );
    case "bar":
      return (
        `<c:plotArea>` +
        `<c:layout/>` +
        `<c:barChart>` +
        `<c:barDir val="bar"/>` +
        `<c:grouping val="clustered"/>` +
        `<c:varyColors val="0"/>` +
        seriesXml +
        `<c:axId val="111111111"/>` +
        `<c:axId val="222222222"/>` +
        `</c:barChart>` +
        defaultCatAxis("111111111", "222222222") +
        defaultValAxis("222222222", "111111111") +
        `</c:plotArea>`
      );
    case "line":
      return (
        `<c:plotArea>` +
        `<c:layout/>` +
        `<c:lineChart>` +
        `<c:grouping val="standard"/>` +
        `<c:varyColors val="0"/>` +
        seriesXml +
        `<c:marker val="1"/>` +
        `<c:axId val="111111111"/>` +
        `<c:axId val="222222222"/>` +
        `</c:lineChart>` +
        defaultCatAxis("111111111", "222222222") +
        defaultValAxis("222222222", "111111111") +
        `</c:plotArea>`
      );
    case "pie":
      return (
        `<c:plotArea>` +
        `<c:layout/>` +
        `<c:pieChart>` +
        `<c:varyColors val="1"/>` +
        seriesXml +
        `<c:firstSliceAng val="0"/>` +
        `</c:pieChart>` +
        `</c:plotArea>`
      );
    default: {
      const _exhaustive: never = kind;
      void _exhaustive;
      return `<c:plotArea><c:layout/></c:plotArea>`;
    }
  }
}

function buildSeries(kind: ChartKind, s: SeriesSpec): string {
  const titleXml = s.titleRef ? `<c:tx><c:strRef><c:f>${escapeXml(s.titleRef)}</c:f></c:strRef></c:tx>` : "";
  const catXml = s.categoriesRef
    ? `<c:cat><c:strRef><c:f>${escapeXml(s.categoriesRef)}</c:f></c:strRef></c:cat>`
    : "";
  const valTag = kind === "pie" ? "c:val" : "c:val";
  const valXml = `<${valTag}><c:numRef><c:f>${escapeXml(s.valuesRange)}</c:f></c:numRef></${valTag}>`;
  // pie charts don't have a meaningful series order axis, but having
  // <c:order> + <c:idx> keeps Excel happy across all chart types.
  return (
    `<c:ser>` +
    `<c:idx val="${s.idx}"/>` +
    `<c:order val="${s.idx}"/>` +
    titleXml +
    (kind === "line" ? `<c:smooth val="0"/>` : "") +
    catXml +
    valXml +
    `</c:ser>`
  );
}

function defaultCatAxis(axId: string, crossAxId: string): string {
  return (
    `<c:catAx>` +
    `<c:axId val="${axId}"/>` +
    `<c:scaling><c:orientation val="minMax"/></c:scaling>` +
    `<c:delete val="0"/>` +
    `<c:axPos val="b"/>` +
    `<c:crossAx val="${crossAxId}"/>` +
    `<c:crosses val="autoZero"/>` +
    `<c:auto val="1"/>` +
    `<c:lblAlgn val="ctr"/>` +
    `<c:lblOffset val="100"/>` +
    `<c:noMultiLvlLbl val="0"/>` +
    `</c:catAx>`
  );
}

function defaultValAxis(axId: string, crossAxId: string): string {
  return (
    `<c:valAx>` +
    `<c:axId val="${axId}"/>` +
    `<c:scaling><c:orientation val="minMax"/></c:scaling>` +
    `<c:delete val="0"/>` +
    `<c:axPos val="l"/>` +
    `<c:crossAx val="${crossAxId}"/>` +
    `<c:crosses val="autoZero"/>` +
    `<c:crossBetween val="between"/>` +
    `</c:valAx>`
  );
}

function emptyChart(kind: ChartKind): string {
  return `<c:chart><c:autoTitleDeleted val="1"/>${buildPlotArea(kind, [])}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>`;
}

function wrapChartSpace(body: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<c:chartSpace xmlns:c="${DRAWINGML_CHART_NS}" xmlns:a="${DRAWINGML_NS}" xmlns:r="${RELS_OFFICE_DOC_NS}">` +
    `<c:roundedCorners val="0"/>` +
    body +
    `</c:chartSpace>`
  );
}

/**
 * Find an unused `xl/charts/chartN.xml` slot in the container. Stable
 * across edits; brand-new charts pick the lowest free integer so old
 * absolute-path references in opaque XML aren't shadowed by accident.
 */
export function mintChartPartPath(
  container: ooxml.OoxmlContainer,
  alreadyMinted: ReadonlySet<string>
): string {
  let i = 1;
  // Walk past both committed parts and parts we're about to write
  // in this same serialize pass.
  while (container.has(`xl/charts/chart${i}.xml`) || alreadyMinted.has(`xl/charts/chart${i}.xml`)) {
    i++;
  }
  return `xl/charts/chart${i}.xml`;
}

/* -------------------------------------------------------------------- */
/* Small parsing + escaping helpers                                      */
/* -------------------------------------------------------------------- */

interface ParsedRange {
  readonly sheet: string;
  readonly startRow: number;
  readonly startCol: number;
  readonly endRow: number;
  readonly endCol: number;
}

/**
 * Parse a `Sheet1!A1:C5` (or `'Has Spaces'!A1:C5`) reference into
 * 0-based row/col coordinates plus the original sheet name (already
 * quoted if needed when re-emitted via {@link quoteSheetName}).
 *
 * Unqualified ranges (no `Sheet!` prefix) are not accepted because
 * DrawingML chart references always need a sheet qualifier — Excel
 * tolerates the omission only when the chart part lives next to a
 * single-sheet workbook, which is brittle in practice.
 */
function parseQualifiedRange(ref: string): ParsedRange | null {
  const bang = ref.indexOf("!");
  if (bang === -1) return null;
  let sheet = ref.slice(0, bang);
  if (sheet.startsWith("'") && sheet.endsWith("'")) {
    sheet = sheet.slice(1, -1).replace(/''/g, "'");
  }
  const range = ref.slice(bang + 1);
  const [startRef, endRef] = range.split(":");
  if (!startRef) return null;
  const start = parseA1(startRef);
  const end = endRef ? parseA1(endRef) : start;
  if (!start || !end) return null;
  return {
    sheet: quoteSheetName(sheet),
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

function parseA1(ref: string): { row: number; col: number } | null {
  // Permit leading `$` anchors — they're meaningless for chart refs
  // (no fill semantics inside the chart part) but a user might type
  // `$A$1:$C$5` and expect it to work.
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref.toUpperCase());
  if (!m) return null;
  const colLetters = m[1]!;
  const row = Number(m[2]!) - 1;
  let col = 0;
  for (const ch of colLetters) col = col * 26 + (ch.charCodeAt(0) - "A".charCodeAt(0) + 1);
  return { row, col: col - 1 };
}

function a1(row: number, col: number): string {
  let n = col + 1;
  let letters = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    letters = String.fromCharCode("A".charCodeAt(0) + r) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return `${letters}${row + 1}`;
}

/**
 * Re-quote a sheet name following Excel conventions: wrap in single
 * quotes when the name contains anything other than the unreserved
 * set, and double up any embedded single quotes.
 */
function quoteSheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
