import type { CommandHandler, HandlerContext } from "@officeai/core";
import type {
  PptxSnapshot,
  TableCell,
  TableRow,
  TableShape,
  TextBody,
  TextParagraph,
  TextRun,
} from "../model/types.js";
import {
  buildDiff,
  evolveSnapshot,
  findShapeInSlide,
  findSlide,
  isTableShape,
  makeError,
  replaceShape,
  withSlide,
} from "./helpers.js";
import type {
  TableAddColumnPayload,
  TableAddRowPayload,
  TableDeleteColumnPayload,
  TableDeleteRowPayload,
  TableSetCellTextPayload,
} from "./payloads.js";

// ─── pptx:table-set-cell-text ────────────────────────────────────────────

export const tableSetCellTextHandler: CommandHandler<TableSetCellTextPayload, PptxSnapshot> = {
  type: "pptx:table-set-cell-text",
  apply(snapshot, payload, ctx) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (!isTableShape(shape)) {
      throw makeError("not-applicable", `shape ${payload.shapeId} is not a table`);
    }
    if (payload.row < 0 || payload.row >= shape.rows.length) {
      throw makeError(
        "unknown-target",
        `row ${payload.row} out of range (0..${shape.rows.length})`
      );
    }
    const row = shape.rows[payload.row]!;
    if (payload.column < 0 || payload.column >= row.cells.length) {
      throw makeError(
        "unknown-target",
        `column ${payload.column} out of range (0..${row.cells.length})`
      );
    }

    const oldCell = row.cells[payload.column]!;
    const newTxBody = textBodyFromPlainText(payload.text, ctx, oldCell.txBody);
    const newCell: TableCell = { ...oldCell, txBody: newTxBody };
    const newRow: TableRow = {
      ...row,
      cells: row.cells.map((c, i) => (i === payload.column ? newCell : c)),
    };
    const newTable: TableShape = {
      ...shape,
      rows: shape.rows.map((r, i) => (i === payload.row ? newRow : r)),
    };

    const root = withSlide(snapshot.root, sIdx, (s) => ({
      ...s,
      shapes: replaceShape(s.shapes, path, newTable),
    }));
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: oldCell.id,
        path: ["slides", sIdx, "shapes", ...path, "rows", payload.row, "cells", payload.column, "txBody"],
        field: "text",
        summary: `+${JSON.stringify(payload.text)}`,
      }),
    };
  },
};

// ─── pptx:table-add-row ──────────────────────────────────────────────────

export const tableAddRowHandler: CommandHandler<TableAddRowPayload, PptxSnapshot> = {
  type: "pptx:table-add-row",
  apply(snapshot, payload, ctx) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (!isTableShape(shape)) {
      throw makeError("not-applicable", `shape ${payload.shapeId} is not a table`);
    }

    const at = payload.at ?? shape.rows.length;
    if (at < 0 || at > shape.rows.length) {
      throw makeError("invalid-position", `at ${at} out of range (0..${shape.rows.length})`);
    }
    const height = payload.height ?? medianRowHeight(shape.rows);
    const newRow: TableRow = {
      id: ctx.mintNodeId(),
      height,
      trAttrs: {},
      cells: shape.columnWidths.map(() => buildEmptyCell(ctx)),
    };

    const newRows = [...shape.rows];
    newRows.splice(at, 0, newRow);
    const newTable: TableShape = { ...shape, rows: newRows };

    const root = withSlide(snapshot.root, sIdx, (s) => ({
      ...s,
      shapes: replaceShape(s.shapes, path, newTable),
    }));
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: newRow.id,
        path: ["slides", sIdx, "shapes", ...path, "rows", at],
        summary: "row",
      }),
    };
  },
};

// ─── pptx:table-delete-row ───────────────────────────────────────────────

export const tableDeleteRowHandler: CommandHandler<TableDeleteRowPayload, PptxSnapshot> = {
  type: "pptx:table-delete-row",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (!isTableShape(shape)) {
      throw makeError("not-applicable", `shape ${payload.shapeId} is not a table`);
    }
    if (shape.rows.length <= 1) {
      throw makeError("invalid-payload", "cannot delete last row of a table");
    }
    if (payload.row < 0 || payload.row >= shape.rows.length) {
      throw makeError(
        "unknown-target",
        `row ${payload.row} out of range (0..${shape.rows.length})`
      );
    }
    const removed = shape.rows[payload.row]!;
    const newRows = shape.rows.filter((_, i) => i !== payload.row);
    const newTable: TableShape = { ...shape, rows: newRows };

    const root = withSlide(snapshot.root, sIdx, (s) => ({
      ...s,
      shapes: replaceShape(s.shapes, path, newTable),
    }));
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-deleted",
        nodeId: removed.id,
        path: ["slides", sIdx, "shapes", ...path, "rows", payload.row],
        summary: "row",
      }),
    };
  },
};

// ─── pptx:table-add-column ───────────────────────────────────────────────

export const tableAddColumnHandler: CommandHandler<TableAddColumnPayload, PptxSnapshot> = {
  type: "pptx:table-add-column",
  apply(snapshot, payload, ctx) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (!isTableShape(shape)) {
      throw makeError("not-applicable", `shape ${payload.shapeId} is not a table`);
    }
    const at = payload.at ?? shape.columnWidths.length;
    if (at < 0 || at > shape.columnWidths.length) {
      throw makeError(
        "invalid-position",
        `at ${at} out of range (0..${shape.columnWidths.length})`
      );
    }
    const oldCount = shape.columnWidths.length;
    const totalWidth = shape.columnWidths.reduce((a, b) => a + b, 0);
    const width =
      payload.width ?? (oldCount > 0 ? Math.round(totalWidth / oldCount) : 1000000);

    const newColumnWidths = [...shape.columnWidths];
    newColumnWidths.splice(at, 0, width);
    const newRows: TableRow[] = shape.rows.map((row) => {
      const cells = [...row.cells];
      cells.splice(at, 0, buildEmptyCell(ctx));
      return { ...row, cells };
    });
    const newTable: TableShape = {
      ...shape,
      columnWidths: newColumnWidths,
      rows: newRows,
    };

    const root = withSlide(snapshot.root, sIdx, (s) => ({
      ...s,
      shapes: replaceShape(s.shapes, path, newTable),
    }));
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: newTable.id,
        path: ["slides", sIdx, "shapes", ...path, "columnWidths", at],
        summary: "column",
      }),
    };
  },
};

// ─── pptx:table-delete-column ────────────────────────────────────────────

export const tableDeleteColumnHandler: CommandHandler<TableDeleteColumnPayload, PptxSnapshot> = {
  type: "pptx:table-delete-column",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (!isTableShape(shape)) {
      throw makeError("not-applicable", `shape ${payload.shapeId} is not a table`);
    }
    if (shape.columnWidths.length <= 1) {
      throw makeError("invalid-payload", "cannot delete last column of a table");
    }
    if (payload.column < 0 || payload.column >= shape.columnWidths.length) {
      throw makeError(
        "unknown-target",
        `column ${payload.column} out of range (0..${shape.columnWidths.length})`
      );
    }
    const c = payload.column;
    const newColumnWidths = shape.columnWidths.filter((_, i) => i !== c);
    const newRows: TableRow[] = shape.rows.map((row) => ({
      ...row,
      cells: row.cells.filter((_, i) => i !== c),
    }));
    const newTable: TableShape = {
      ...shape,
      columnWidths: newColumnWidths,
      rows: newRows,
    };

    const root = withSlide(snapshot.root, sIdx, (s) => ({
      ...s,
      shapes: replaceShape(s.shapes, path, newTable),
    }));
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-deleted",
        nodeId: newTable.id,
        path: ["slides", sIdx, "shapes", ...path, "columnWidths", c],
        summary: "column",
      }),
    };
  },
};

// ─── helpers ─────────────────────────────────────────────────────────────

function buildEmptyCell(ctx: HandlerContext): TableCell {
  const run: TextRun = { id: ctx.mintNodeId(), properties: {}, text: "" };
  const paragraph: TextParagraph = {
    id: ctx.mintNodeId(),
    properties: {},
    runs: [run],
  };
  return {
    id: ctx.mintNodeId(),
    txBody: { paragraphs: [paragraph] },
    tcAttrs: {},
  };
}

function textBodyFromPlainText(
  text: string,
  ctx: HandlerContext,
  inherit: TextBody
): TextBody {
  const lines = text.length === 0 ? [""] : text.split("\n");
  const inheritedParaProps = inherit.paragraphs[0]?.properties ?? {};
  const inheritedRunProps = inherit.paragraphs[0]?.runs[0]?.properties ?? {};
  const paragraphs: TextParagraph[] = lines.map((line) => {
    const run: TextRun = {
      id: ctx.mintNodeId(),
      properties: { ...inheritedRunProps },
      text: line,
    };
    return {
      id: ctx.mintNodeId(),
      properties: { ...inheritedParaProps },
      runs: [run],
    };
  });
  return {
    ...inherit,
    paragraphs,
  };
}

function medianRowHeight(rows: ReadonlyArray<TableRow>): number {
  if (rows.length === 0) return 0;
  const sorted = [...rows.map((r) => r.height)].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 1
    ? sorted[mid]!
    : Math.round(((sorted[mid - 1] ?? 0) + (sorted[mid] ?? 0)) / 2);
}
