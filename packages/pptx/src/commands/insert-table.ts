import { ooxml, type CommandHandler, type HandlerContext } from "@officeai/core";
import type {
  OpaquePart,
  OpaqueXml,
  PptxSnapshot,
  Slide,
  TableCell,
  TableRow,
  TableShape,
  TextBody,
  TextParagraph,
  TextRun,
  TextRunProperties,
} from "../model/types.js";
import { attrOf, elementEntries, findElementEntry } from "../parser/xml-helpers.js";
import { buildDiff, evolveSnapshot, findSlide, makeError, maxCNvPrId } from "./helpers.js";
import type { InsertTablePayload } from "./payloads.js";

/**
 * `<a:graphicData @uri>` value PowerPoint uses to identify a table
 * graphic frame. Must match `parseGraphicFrameTable` in the parser
 * (and the legacy table commands), otherwise a freshly-inserted
 * table won't round-trip back as a typed `TableShape`.
 */
const TABLE_GRAPHIC_DATA_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";

/** Default frame: 8" × 4" centred on the slide. */
const DEFAULT_WIDTH_EMU = 7_315_200;
const DEFAULT_HEIGHT_EMU = 3_657_600;

/** Calibri 18pt fallback when the master doesn't expose a body style. */
const FALLBACK_FONT_FAMILY = "Calibri";
const FALLBACK_FONT_SIZE_HUNDREDTHS = 1800;

export const insertTableHandler: CommandHandler<InsertTablePayload, PptxSnapshot> = {
  type: "pptx:insert-table",
  apply(snapshot, payload, ctx) {
    if (!Number.isInteger(payload.rows) || payload.rows <= 0) {
      throw makeError("invalid-payload", `rows must be a positive integer (got ${payload.rows})`);
    }
    if (!Number.isInteger(payload.cols) || payload.cols <= 0) {
      throw makeError("invalid-payload", `cols must be a positive integer (got ${payload.cols})`);
    }
    if (payload.widthEmu !== undefined && payload.widthEmu <= 0) {
      throw makeError("invalid-payload", "widthEmu must be > 0");
    }
    if (payload.heightEmu !== undefined && payload.heightEmu <= 0) {
      throw makeError("invalid-payload", "heightEmu must be > 0");
    }

    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);

    const slideSize = snapshot.root.slideSize;
    const widthEmu = Math.round(payload.widthEmu ?? Math.min(DEFAULT_WIDTH_EMU, slideSize.cxEmu));
    const heightEmu = Math.round(payload.heightEmu ?? Math.min(DEFAULT_HEIGHT_EMU, slideSize.cyEmu));
    const xEmu = Math.round(payload.xEmu ?? Math.max(0, (slideSize.cxEmu - widthEmu) / 2));
    const yEmu = Math.round(payload.yEmu ?? Math.max(0, (slideSize.cyEmu - heightEmu) / 2));

    const cNvPrId = maxCNvPrId(slide.shapes) + 1;
    const name = payload.name ?? `Table ${cNvPrId}`;

    const columnWidths = splitEvenly(widthEmu, payload.cols);
    const rowHeights = splitEvenly(heightEmu, payload.rows);

    const defaultRunProps = readDefaultBodyRunProps(snapshot);

    const rows: TableRow[] = [];
    for (let r = 0; r < payload.rows; r++) {
      const cells: TableCell[] = [];
      for (let c = 0; c < payload.cols; c++) {
        const text = payload.cells?.[r]?.[c] ?? "";
        cells.push(buildCell(ctx, text, defaultRunProps));
      }
      rows.push({
        id: ctx.mintNodeId(),
        height: rowHeights[r]!,
        cells,
        trAttrs: {},
      });
    }

    const table: TableShape = {
      kind: "table",
      id: ctx.mintNodeId(),
      cNvPrId,
      name,
      position: { xEmu, yEmu },
      size: { cxEmu: widthEmu, cyEmu: heightEmu },
      columnWidths,
      rows,
      tblPrRaw: defaultTblPr(),
      nvGraphicFramePrTail: defaultNvGraphicFramePrTail(),
      graphicDataUri: TABLE_GRAPHIC_DATA_URI,
    };

    const newSlide: Slide = { ...slide, shapes: [...slide.shapes, table] };
    const root = {
      ...snapshot.root,
      slides: snapshot.root.slides.map((s, i) => (i === sIdx ? newSlide : s)),
    };
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: table.id,
        path: ["slides", sIdx, "shapes", newSlide.shapes.length - 1],
        summary: `table ${payload.rows}×${payload.cols}`,
      }),
    };
  },
};

// ─── Helpers ──────────────────────────────────────────────────────────────

function splitEvenly(total: number, parts: number): number[] {
  const base = Math.floor(total / parts);
  const remainder = total - base * parts;
  const out: number[] = new Array(parts).fill(base);
  // Spread the rounding remainder across the leading entries so the
  // sum is exactly `total` (PowerPoint validates that gridCol widths
  // sum to the frame extent on some templates).
  for (let i = 0; i < remainder; i++) out[i] += 1;
  return out;
}

function buildCell(ctx: HandlerContext, text: string, defaultRunProps: TextRunProperties): TableCell {
  const lines = text.length === 0 ? [""] : text.split("\n");
  const paragraphs: TextParagraph[] = lines.map((line) => {
    const run: TextRun = {
      id: ctx.mintNodeId(),
      properties: { ...defaultRunProps },
      text: line,
    };
    return { id: ctx.mintNodeId(), properties: {}, runs: [run] };
  });
  const txBody: TextBody = {
    bodyPrRaw: defaultBodyPr(),
    paragraphs,
  };
  return {
    id: ctx.mintNodeId(),
    txBody,
    tcPrRaw: defaultTcPr(),
    tcAttrs: {},
  };
}

/**
 * Walk the first slide master's `<p:txStyles>/<p:bodyStyle>/<a:lvl1pPr>/<a:defRPr>`
 * and lift the typeface + size into a `TextRunProperties` bag we can
 * apply to fresh cells. Returns Calibri/18pt when nothing is found
 * (no master, parser couldn't capture txStyles, …) so the resulting
 * cells always render with concrete formatting in PowerPoint
 * regardless of the deck's master state.
 */
function readDefaultBodyRunProps(snapshot: PptxSnapshot): TextRunProperties {
  const fallback: TextRunProperties = {
    fontFamily: FALLBACK_FONT_FAMILY,
    fontSizeHundredths: FALLBACK_FONT_SIZE_HUNDREDTHS,
  };
  const master = firstMaster(snapshot);
  if (!master) return fallback;
  const defRPr = findDefBodyRPr(master.raw);
  if (!defRPr) return fallback;
  const props: { -readonly [K in keyof TextRunProperties]: TextRunProperties[K] } = { ...fallback };
  const sz = attrOf(defRPr, "sz");
  if (sz !== undefined) {
    const n = Number(sz);
    if (Number.isFinite(n) && n > 0) props.fontSizeHundredths = n;
  }
  const subtree = (defRPr["a:defRPr"] as unknown[] | undefined) ?? [];
  const latin = findElementEntry(subtree, "a:latin");
  if (latin) {
    const typeface = attrOf(latin, "typeface");
    if (typeface) props.fontFamily = typeface;
  }
  return props;
}

function firstMaster(snapshot: PptxSnapshot): OpaquePart | null {
  const it = snapshot.root.masters.values().next();
  return it.done ? null : it.value;
}

function findDefBodyRPr(master: OpaqueXml): Record<string, unknown> | null {
  if (master.tag !== "p:sldMaster") return null;
  const txStyles = findInOpaqueChildren(master.subtree, "p:txStyles");
  if (!txStyles) return null;
  const bodyStyle = findInElementChildren(txStyles, "p:bodyStyle");
  if (!bodyStyle) return null;
  const lvl1 = findInElementChildren(bodyStyle, "a:lvl1pPr");
  if (!lvl1) return null;
  return findInElementChildren(lvl1, "a:defRPr");
}

function findInOpaqueChildren(
  children: ReadonlyArray<unknown>,
  tag: string
): Record<string, unknown> | null {
  for (const c of elementEntries(children)) {
    if (ooxml.getTag(c) === tag) return c;
  }
  return null;
}

function findInElementChildren(parent: Record<string, unknown>, tag: string): Record<string, unknown> | null {
  const parentTag = ooxml.getTag(parent);
  const children = (parent[parentTag] as unknown[] | undefined) ?? [];
  return findInOpaqueChildren(children, tag);
}

function defaultNvGraphicFramePrTail(): OpaqueXml[] {
  return [
    {
      tag: "p:cNvPr",
      attrs: { id: "0", name: "" },
      rawAttrs: { "@_id": "0", "@_name": "" },
      subtree: [],
    },
    {
      tag: "p:cNvGraphicFramePr",
      attrs: {},
      rawAttrs: {},
      subtree: [{ "a:graphicFrameLocks": [], ":@": { "@_noGrp": "1" } }],
    },
    { tag: "p:nvPr", attrs: {}, rawAttrs: {}, subtree: [] },
  ];
}

function defaultTblPr(): OpaqueXml {
  return {
    tag: "a:tblPr",
    attrs: { firstRow: "1", bandRow: "1" },
    rawAttrs: { "@_firstRow": "1", "@_bandRow": "1" },
    subtree: [],
  };
}

function defaultBodyPr(): OpaqueXml {
  return {
    tag: "a:bodyPr",
    attrs: {},
    rawAttrs: {},
    subtree: [],
  };
}

function defaultTcPr(): OpaqueXml {
  return {
    tag: "a:tcPr",
    attrs: {},
    rawAttrs: {},
    subtree: [],
  };
}
