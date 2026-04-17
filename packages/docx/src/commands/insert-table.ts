import { CommandError, type CommandHandler, type IdMinter } from "@officeai/core";
import type {
  BlockNode,
  DocxDocument,
  DocxSnapshot,
  Paragraph,
  Table,
  TableCell,
  TableGridCol,
  TableProperties,
  TableRow,
} from "../model/types.js";
import { buildDiff, evolveSnapshot, insertBlock } from "./helpers.js";
import type { InsertTablePayload } from "./payloads.js";

/**
 * Insert a fresh `rows` × `cols` table at the given body offset.
 *
 * The new table is fully typed (no `raw` blob) so the serializer regenerates
 * it from the model. Default cell content per the brief: one empty
 * paragraph. `columnWidths` (twips) drives per-column `<w:gridCol>` widths
 * AND per-cell `<w:tcW>` widths, so the rendered table matches what the
 * caller asked for. When omitted, the grid carries `<w:gridCol>` entries
 * with no `w:w` attribute (Word picks a default), and cells get no `tcW`.
 */
export const insertTableHandler: CommandHandler<InsertTablePayload, DocxSnapshot> = {
  type: "docx:insert-table",
  apply(snapshot, payload, ctx) {
    const { at, rows, cols, columnWidths, properties } = payload;
    if (!Number.isInteger(rows) || rows < 1) {
      throw new CommandError("invalid-payload", `rows must be a positive integer (got ${rows})`);
    }
    if (!Number.isInteger(cols) || cols < 1) {
      throw new CommandError("invalid-payload", `cols must be a positive integer (got ${cols})`);
    }
    if (columnWidths && columnWidths.length !== cols) {
      throw new CommandError(
        "invalid-payload",
        `columnWidths.length (${columnWidths.length}) does not match cols (${cols})`
      );
    }
    const bodyLen = snapshot.root.body.length;
    if (!Number.isInteger(at.paragraph) || at.paragraph < 0 || at.paragraph > bodyLen) {
      throw new CommandError("invalid-position", `body offset ${at.paragraph} out of range [0, ${bodyLen}]`);
    }

    const table = makeBlankTable({
      rows,
      cols,
      columnWidths,
      properties,
      mintNodeId: ctx.mintNodeId,
    });

    const insertedAt = at.paragraph;
    const nextDoc =
      insertedAt === bodyLen
        ? appendBlock(snapshot.root, table)
        : insertBlock(snapshot.root, insertedAt, table);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: table.id,
        path: ["body", insertedAt],
        summary: `+table ${rows}×${cols}`,
      }),
    };
  },
};

function appendBlock(doc: DocxDocument, block: BlockNode): DocxDocument {
  return { ...doc, body: [...doc.body, block] };
}

/* ── shared builders / lookup helpers used by every typed-table handler ─── */

export interface MakeBlankTableOptions {
  readonly rows: number;
  readonly cols: number;
  readonly columnWidths?: ReadonlyArray<number>;
  readonly properties?: Partial<TableProperties>;
  readonly mintNodeId: IdMinter;
}

/**
 * Build a brand-new typed `Table` with an empty paragraph in every cell.
 * Always returns a table without `raw` so the serializer regenerates it.
 */
export function makeBlankTable(opts: MakeBlankTableOptions): Table {
  const { rows, cols, columnWidths, properties, mintNodeId } = opts;
  const grid: TableGridCol[] = [];
  for (let c = 0; c < cols; c++) {
    if (columnWidths && columnWidths[c] !== undefined) grid.push({ w: columnWidths[c] });
    else grid.push({});
  }
  const tableRows: TableRow[] = [];
  for (let r = 0; r < rows; r++) {
    tableRows.push(makeBlankRow(cols, grid, mintNodeId));
  }
  return {
    kind: "table",
    id: mintNodeId(),
    properties: { ...(properties ?? {}) },
    grid,
    rows: tableRows,
  };
}

/**
 * Build a fresh row of `cols` cells. Every cell carries an empty paragraph
 * and inherits its `<w:tcW>` width from the grid (when set), so insertions
 * preserve column widths without the caller having to recompute them.
 */
export function makeBlankRow(
  cols: number,
  grid: ReadonlyArray<TableGridCol>,
  mintNodeId: IdMinter
): TableRow {
  const cells: TableCell[] = [];
  for (let c = 0; c < cols; c++) {
    cells.push(makeBlankCell(grid[c]?.w, mintNodeId));
  }
  return {
    kind: "table-row",
    id: mintNodeId(),
    properties: {},
    cells,
  };
}

export function makeBlankCell(width: number | undefined, mintNodeId: IdMinter): TableCell {
  const para: Paragraph = {
    kind: "paragraph",
    id: mintNodeId(),
    properties: {},
    children: [{ kind: "run", id: mintNodeId(), properties: {}, children: [] }],
  };
  return {
    kind: "table-cell",
    id: mintNodeId(),
    properties: width !== undefined ? { tcW: { value: width, type: "dxa" } } : {},
    body: [para],
  };
}

/**
 * Drop the `raw` cache so the serializer regenerates the table from the
 * typed model. Must be called whenever a mutating command produces a new
 * table.
 */
export function withoutRaw(table: Table): Table {
  if (table.raw === undefined) return table;
  const { raw: _ignored, ...rest } = table;
  void _ignored;
  return rest;
}

/* ── Recursive table lookup ──────────────────────────────────────────────── */

export type TableUpdater = (next: Table) => DocxDocument;

export interface TableLocation {
  readonly table: Table;
  /** Ordered list of ancestor table ids (outermost first). */
  readonly ancestorIds: ReadonlyArray<string>;
  /**
   * Replace the located table with a new one and return the updated
   * `DocxDocument`. The replacement is performed at the same depth, dropping
   * `raw` on every ancestor table along the way (since their typed shape
   * also changed).
   */
  readonly replace: TableUpdater;
  /** Body index of the outermost ancestor table. Useful for diff paths. */
  readonly bodyIndex: number;
}

/**
 * Find a table by id anywhere in the body (including inside cells of other
 * tables). Returns `null` if no match. The returned `replace` closure
 * rebuilds every enclosing table (dropping their `raw` cache) so the change
 * is visible to the serializer at every level.
 */
export function findTable(doc: DocxDocument, tableId: string): TableLocation | null {
  for (let i = 0; i < doc.body.length; i++) {
    const block = doc.body[i];
    if (block.kind !== "table") continue;
    const found = findTableInTable(block, tableId, []);
    if (found) {
      const bodyIndex = i;
      const replace = (next: Table): DocxDocument => {
        const newBlock = found.rebuild(next);
        const body = doc.body.slice();
        body[bodyIndex] = newBlock;
        return { ...doc, body };
      };
      return { table: found.table, ancestorIds: found.ancestorIds, replace, bodyIndex };
    }
  }
  return null;
}

interface InnerLocation {
  readonly table: Table;
  readonly ancestorIds: ReadonlyArray<string>;
  /** Returns the (possibly outermost) table containing the replacement. */
  readonly rebuild: (next: Table) => Table;
}

function findTableInTable(table: Table, tableId: string, ancestors: string[]): InnerLocation | null {
  if (table.id === tableId) {
    return {
      table,
      ancestorIds: ancestors.slice(),
      rebuild: (next) => next,
    };
  }
  for (let r = 0; r < table.rows.length; r++) {
    const row = table.rows[r];
    for (let c = 0; c < row.cells.length; c++) {
      const cell = row.cells[c];
      for (let b = 0; b < cell.body.length; b++) {
        const block = cell.body[b];
        if (block.kind !== "table") continue;
        const found = findTableInTable(block, tableId, [...ancestors, table.id]);
        if (!found) continue;
        const rebuild = (next: Table): Table => {
          const inner = found.rebuild(next);
          const newBody = cell.body.slice();
          newBody[b] = inner;
          const newCell: TableCell = { ...cell, body: newBody };
          const newCells = row.cells.slice();
          newCells[c] = newCell;
          const newRow: TableRow = { ...row, cells: newCells };
          const newRows = table.rows.slice();
          newRows[r] = newRow;
          return withoutRaw({ ...table, rows: newRows });
        };
        return { table: found.table, ancestorIds: found.ancestorIds, rebuild };
      }
    }
  }
  return null;
}
