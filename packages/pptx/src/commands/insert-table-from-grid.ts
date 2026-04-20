import type { CommandHandler } from "@officeai/core";
import type {
  PptxPresentation,
  PptxSnapshot,
  Slide,
  TableCell,
  TableRow,
  TableShape,
  TextBody,
  TextParagraph,
  TextRun,
} from "../model/types.js";
import { evolveSnapshot, findSlide, makeError, maxCNvPrId } from "./helpers.js";
import type { PptxInsertTableFromGridPayload } from "./payloads.js";

const TABLE_GRAPHIC_DATA_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";

/**
 * Author a real `TableShape` on a slide from a 2D grid (typically the
 * result of pasting / dropping a range from xlsx). Mirrors the shape
 * of `pptx:insert-image` for sizing/positioning, but produces a typed
 * `<a:tbl>` graphic frame instead of a text box, so the user can edit
 * cells in place rather than re-flow plain TSV.
 *
 * Sizing rules:
 *   - `cx` (or column total) is split evenly across columns when no
 *     `columnWidthsEmu` is provided.
 *   - `cy` is split evenly across rows when no `rowHeightsEmu` is
 *     provided. Both default to a small ~5×0.5-inch-per-row block.
 */
export const insertTableFromGridHandler: CommandHandler<
  PptxInsertTableFromGridPayload,
  PptxSnapshot
> = {
  type: "pptx:insert-table",
  apply(snapshot, payload, ctx) {
    if (!Array.isArray(payload.data) || payload.data.length === 0) {
      throw makeError("invalid-payload", "data must contain at least one row");
    }
    const cols = payload.data.reduce((m, r) => Math.max(m, r.length), 0);
    if (cols === 0) {
      throw makeError("invalid-payload", "data must contain at least one column");
    }

    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);

    const rows = payload.data.length;
    const totalCx = Math.round(payload.cx ?? 4_572_000); // ~5 inches
    const totalCy = Math.round(payload.cy ?? Math.max(rows, 1) * 457_200); // ~0.5 inch/row

    const columnWidths = resolveColumnWidths(payload.columnWidthsEmu, cols, totalCx);
    const rowHeights = resolveRowHeights(payload.rowHeightsEmu, rows, totalCy);

    const tableRows: TableRow[] = payload.data.map((row, rIdx): TableRow => {
      const cells: TableCell[] = [];
      for (let c = 0; c < cols; c++) {
        const value = row[c];
        cells.push(buildCell(value, ctx.mintNodeId));
      }
      return {
        id: ctx.mintNodeId(),
        height: rowHeights[rIdx]!,
        cells,
        trAttrs: {},
      };
    });

    const cNvPrId = maxCNvPrId(slide.shapes) + 1;
    const table: TableShape = {
      kind: "table",
      id: ctx.mintNodeId(),
      cNvPrId,
      name: payload.name ?? `Table ${cNvPrId}`,
      position: { xEmu: Math.round(payload.x), yEmu: Math.round(payload.y) },
      size: { cxEmu: totalCx, cyEmu: totalCy },
      columnWidths,
      rows: tableRows,
      nvGraphicFramePrTail: [],
      graphicDataUri: TABLE_GRAPHIC_DATA_URI,
    };

    const newSlide: Slide = { ...slide, shapes: [...slide.shapes, table] };
    const newSlides = [...snapshot.root.slides];
    newSlides[sIdx] = newSlide;

    const root: PptxPresentation = { ...snapshot.root, slides: newSlides };
    const next = evolveSnapshot(snapshot, root, {
      slides: [slide.partPath],
    });

    return {
      next,
      diff: {
        format: "pptx",
        fromRevision: snapshot.revision,
        toRevision: next.revision,
        changes: [
          {
            kind: "node-inserted",
            nodeId: table.id,
            path: [
              "slides",
              sIdx,
              "shapes",
              newSlide.shapes.length - 1,
            ] as ReadonlyArray<string | number>,
            summary: `+table (${rows}×${cols})`,
          },
        ],
      },
    };
  },
};

// ─── helpers ──────────────────────────────────────────────────────────────

function resolveColumnWidths(
  override: ReadonlyArray<number> | undefined,
  cols: number,
  totalCx: number
): number[] {
  if (override && override.length === cols) {
    return override.map((w) => Math.round(w));
  }
  const each = Math.floor(totalCx / Math.max(cols, 1));
  const out: number[] = new Array(cols).fill(each);
  // Distribute rounding remainder onto the last column so the
  // grid-total matches the table's `cx` exactly.
  out[cols - 1] = totalCx - each * (cols - 1);
  return out;
}

function resolveRowHeights(
  override: ReadonlyArray<number> | undefined,
  rows: number,
  totalCy: number
): number[] {
  if (override && override.length === rows) {
    return override.map((h) => Math.round(h));
  }
  const each = Math.floor(totalCy / Math.max(rows, 1));
  const out: number[] = new Array(rows).fill(each);
  out[rows - 1] = totalCy - each * (rows - 1);
  return out;
}

function buildCell(
  value: string | number | null | undefined,
  mintNodeId: () => string
): TableCell {
  const text = stringifyCellValue(value);
  const txBody = textBodyFromString(text, mintNodeId);
  return {
    id: mintNodeId(),
    txBody,
    tcAttrs: {},
  };
}

function stringifyCellValue(value: string | number | null | undefined): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return "";
    return String(value);
  }
  return value;
}

function textBodyFromString(text: string, mintNodeId: () => string): TextBody {
  const run: TextRun = {
    id: mintNodeId(),
    properties: {},
    text,
  };
  const paragraph: TextParagraph = {
    id: mintNodeId(),
    properties: {},
    runs: [run],
  };
  return { paragraphs: [paragraph] };
}
