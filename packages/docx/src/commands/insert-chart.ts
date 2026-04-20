import { CommandError, type CommandHandler, type IdMinter } from "@officeai/core";
import type {
  ChartDrawing,
  ChartPart,
  ChartSeries,
  ChartType,
  DocxDocument,
  DocxSnapshot,
  InlineNode,
  Paragraph,
  Relationship,
  Run,
  RunChild,
} from "../model/types.js";
import { buildDiffMulti, evolveSnapshot } from "./helpers.js";
import { mintDocPrId } from "./insert-image.js";
import type { InsertChartPayload } from "./payloads.js";

/**
 * Insert a typed chart at the given paragraph position.
 *
 * Steps:
 *   1. Validate the payload (positive dimensions, ≥ 1 category and
 *      series, all series share the same length as `categories`).
 *   2. Mint a fresh `word/charts/chartN.xml` part path and wire a new
 *      `chart`-type relationship from `word/document.xml.rels`.
 *   3. Build a typed `ChartPart` carrying the categories, series, type,
 *      title — the serializer reads this to (re)build the chart XML
 *      and the embedded xlsx workbook.
 *   4. Splice a typed `ChartDrawing` leaf into the targeted paragraph.
 *   5. Mark dirty: body (the spliced run), charts (the new chart part),
 *      relationships (document.xml.rels), contentTypes (chart MIME +
 *      embedded xlsx MIME — applied by the serializer when it emits the
 *      part).
 */
export const insertChartHandler: CommandHandler<InsertChartPayload, DocxSnapshot> = {
  type: "docx:insert-chart",
  apply(snapshot, payload, ctx) {
    validatePayload(payload);

    const bodyLen = snapshot.root.body.length;
    if (
      !Number.isInteger(payload.at.paragraph) ||
      payload.at.paragraph < 0 ||
      payload.at.paragraph >= bodyLen
    ) {
      throw new CommandError(
        "invalid-position",
        `paragraph index ${payload.at.paragraph} out of range [0, ${bodyLen})`
      );
    }
    const target = snapshot.root.body[payload.at.paragraph];
    if (target.kind !== "paragraph") {
      throw new CommandError(
        "not-paragraph",
        `block at body index ${payload.at.paragraph} is not a paragraph (kind=${target.kind})`
      );
    }

    const docRelsKey = "word/document.xml";
    const docRels = snapshot.root.relationships.get(docRelsKey) ?? [];
    const chartPartPath = mintChartPartPath(snapshot.root.charts);
    const relTarget = chartPartPath.startsWith("word/") ? chartPartPath.slice("word/".length) : chartPartPath;
    const relId = mintRelId(docRels);
    const newRel: Relationship = { id: relId, type: CHART_REL_TYPE, target: relTarget };
    const nextRels: ReadonlyArray<Relationship> = [...docRels, newRel];

    const docPrId = mintDocPrId(snapshot.root);
    const widthPx = payload.width ?? 480;
    const heightPx = payload.height ?? 320;
    const cx = pxToEmu(widthPx);
    const cy = pxToEmu(heightPx);

    const chartPart: ChartPart = {
      partPath: chartPartPath,
      contentType: CHART_CONTENT_TYPE,
      chartType: payload.chartType as ChartType,
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
      embeddingSheetName: "Sheet1",
    };

    const drawing: ChartDrawing = {
      kind: "drawing",
      subkind: "chart",
      id: ctx.mintNodeId(),
      relId,
      chartPartPath,
      cx,
      cy,
      docPrId,
      name: payload.name ?? `Chart ${docPrId}`,
      ...(payload.altText !== undefined ? { descr: payload.altText } : {}),
    };

    const updatedParagraph = insertDrawingIntoParagraph(
      target,
      payload.at.run,
      payload.at.offset ?? 0,
      drawing,
      ctx.mintNodeId
    );
    const newBody = snapshot.root.body.slice();
    newBody[payload.at.paragraph] = updatedParagraph;

    const newCharts = new Map(snapshot.root.charts);
    newCharts.set(chartPartPath, chartPart);
    const newRelationships = new Map(snapshot.root.relationships);
    newRelationships.set(docRelsKey, nextRels);

    const nextDoc: DocxDocument = {
      ...snapshot.root,
      body: newBody,
      charts: newCharts,
      relationships: newRelationships,
    };

    const next = evolveSnapshot(snapshot, nextDoc, {
      body: true,
      charts: withAddition(snapshot.dirty.charts, chartPartPath),
      relationships: withAddition(snapshot.dirty.relationships, docRelsKey),
      contentTypes: true,
    });

    return {
      next,
      diff: buildDiffMulti(snapshot.revision, next.revision, [
        {
          kind: "node-inserted",
          nodeId: drawing.id,
          path: ["body", payload.at.paragraph, "chart"],
          summary: `+chart (${payload.chartType}, ${payload.categories.length} cats × ${payload.series.length} ser)`,
        },
        {
          kind: "part-added",
          path: [chartPartPath],
          summary: `+chart ${chartPartPath}`,
        },
      ]),
    };
  },
};

const CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

function validatePayload(p: InsertChartPayload): void {
  if (!Array.isArray(p.categories) || p.categories.length === 0) {
    throw new CommandError("invalid-payload", "categories must contain at least one entry");
  }
  if (!Array.isArray(p.series) || p.series.length === 0) {
    throw new CommandError("invalid-payload", "series must contain at least one entry");
  }
  for (let i = 0; i < p.series.length; i++) {
    const s = p.series[i]!;
    if (!Array.isArray(s.values)) {
      throw new CommandError("invalid-payload", `series[${i}].values must be an array`);
    }
    if (s.values.length !== p.categories.length) {
      throw new CommandError(
        "invalid-payload",
        `series[${i}].values length ${s.values.length} ≠ categories length ${p.categories.length}`
      );
    }
  }
  if (p.width !== undefined && (!Number.isFinite(p.width) || p.width <= 0)) {
    throw new CommandError("invalid-payload", `width must be a positive number (got ${p.width})`);
  }
  if (p.height !== undefined && (!Number.isFinite(p.height) || p.height <= 0)) {
    throw new CommandError("invalid-payload", `height must be a positive number (got ${p.height})`);
  }
}

function mintChartPartPath(charts: ReadonlyMap<string, ChartPart>): string {
  let n = 1;
  while (charts.has(`word/charts/chart${n}.xml`)) n++;
  return `word/charts/chart${n}.xml`;
}

function mintRelId(rels: ReadonlyArray<Relationship>): string {
  const taken = new Set(rels.map((r) => r.id));
  let i = rels.length + 1;
  while (taken.has(`rId${i}`)) i++;
  return `rId${i}`;
}

function pxToEmu(px: number): number {
  return Math.round(px * 9525);
}

function withAddition(prev: ReadonlySet<string>, member: string): ReadonlySet<string> {
  const next = new Set(prev);
  next.add(member);
  return next;
}

/**
 * Splice a new run carrying the chart drawing leaf into a paragraph.
 * Mirrors `insert-image`'s splice helper so user expectations stay
 * consistent across inline media insertions.
 */
function insertDrawingIntoParagraph(
  p: Paragraph,
  runIndex: number | undefined,
  offset: number,
  leaf: ChartDrawing,
  mintNodeId: IdMinter
): Paragraph {
  const drawingRun: Run = {
    kind: "run",
    id: mintNodeId(),
    properties: {},
    children: [leaf as RunChild],
  };

  if (runIndex === undefined) {
    return { ...p, children: [drawingRun, ...p.children] };
  }
  if (runIndex < 0 || runIndex >= p.children.length) {
    return { ...p, children: [...p.children, drawingRun] };
  }
  const target = p.children[runIndex];
  if (target.kind !== "run") {
    const next = p.children.slice();
    next.splice(runIndex, 0, drawingRun);
    return { ...p, children: next };
  }

  const { before, after } = splitRunAtOffset(target, offset, mintNodeId);
  const next: InlineNode[] = [];
  for (let i = 0; i < p.children.length; i++) {
    if (i === runIndex) {
      if (before) next.push(before);
      next.push(drawingRun);
      if (after) next.push(after);
    } else {
      next.push(p.children[i]);
    }
  }
  return { ...p, children: next };
}

interface SplitRun {
  before: Run | null;
  after: Run | null;
}

function splitRunAtOffset(run: Run, offset: number, mintNodeId: IdMinter): SplitRun {
  const beforeChildren: RunChild[] = [];
  const afterChildren: RunChild[] = [];
  let consumed = 0;
  let placed = false;
  for (const c of run.children) {
    if (placed) {
      afterChildren.push(c);
      continue;
    }
    if (c.kind !== "text") {
      beforeChildren.push(c);
      continue;
    }
    const len = c.text.length;
    if (offset >= consumed + len) {
      beforeChildren.push(c);
      consumed += len;
      continue;
    }
    const local = Math.max(0, offset - consumed);
    if (local > 0) beforeChildren.push({ ...c, text: c.text.slice(0, local) });
    if (local < len) {
      afterChildren.push({ ...c, id: mintNodeId(), text: c.text.slice(local) });
    }
    placed = true;
    consumed += len;
  }
  const before: Run | null =
    beforeChildren.length > 0 ? { ...run, id: mintNodeId(), children: beforeChildren } : null;
  const after: Run | null =
    afterChildren.length > 0 ? { ...run, id: mintNodeId(), children: afterChildren } : null;
  return { before, after };
}
