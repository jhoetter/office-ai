import { CommandError, type CommandHandler } from "@officeai/core";
import type {
  DocxDocument,
  DocxSnapshot,
  Shading,
  Table,
  TableCell,
  TableCellProperties,
  TableRow,
  TableRowProperties,
} from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { findTable, withoutRaw } from "./insert-table.js";
import type { TableLocation } from "./insert-table.js";

/**
 * Phase 3e — Table tools: shading + alignment + sizing + simple merges.
 *
 * These all share the same `(tableId, row, column)` selector and route
 * through `findTable` / `withoutRaw` so the serializer drops cached XML
 * and re-emits from the typed model — preserving every untouched table
 * byte-identically through the table `raw` cache.
 */

export interface SetCellShadingPayload {
  readonly tableId: string;
  /** 0-based row index. */
  readonly row: number;
  /** 0-based grid column index (post-merge collapse). */
  readonly column: number;
  /** Hex RGB fill (no leading `#`). Pass `null` to clear shading entirely. */
  readonly fill: string | null;
  /** Optional foreground / pattern colour (hex RGB). */
  readonly color?: string;
  /** Optional OOXML shading pattern val (`clear`, `solid`, `pct25`, …). */
  readonly pattern?: string;
}

export const setCellShadingHandler: CommandHandler<SetCellShadingPayload, DocxSnapshot> = {
  type: "docx:set-cell-shading",
  apply(snapshot, payload) {
    return mutateCell(snapshot, payload, (props) => {
      if (payload.fill === null) {
        const { shd: _shd, ...rest } = props;
        return rest;
      }
      const shd: Shading = {
        fill: normaliseHex(payload.fill),
        ...(payload.color !== undefined ? { color: normaliseHex(payload.color) } : {}),
        ...(payload.pattern !== undefined ? { pattern: payload.pattern } : {}),
      };
      return { ...props, shd };
    }, "shading");
  },
};

export interface SetCellAlignmentPayload {
  readonly tableId: string;
  readonly row: number;
  readonly column: number;
  /** Pass `null` to clear the explicit alignment (Word will fall back to top). */
  readonly vAlign: "top" | "center" | "bottom" | null;
}

export const setCellAlignmentHandler: CommandHandler<SetCellAlignmentPayload, DocxSnapshot> = {
  type: "docx:set-cell-alignment",
  apply(snapshot, payload) {
    return mutateCell(snapshot, payload, (props) => {
      if (payload.vAlign === null) {
        const { vAlign: _v, ...rest } = props;
        return rest;
      }
      return { ...props, vAlign: payload.vAlign };
    }, "alignment");
  },
};

export interface SetRowHeightPayload {
  readonly tableId: string;
  readonly row: number;
  /** Height in twips. Pass `null` to clear (Word falls back to "auto"). */
  readonly heightTwips: number | null;
  /** `auto` (default), `exact`, or `atLeast`. */
  readonly rule?: "auto" | "exact" | "atLeast";
}

export const setRowHeightHandler: CommandHandler<SetRowHeightPayload, DocxSnapshot> = {
  type: "docx:set-row-height",
  apply(snapshot, payload) {
    return mutateRow(snapshot, payload, (props) => {
      if (payload.heightTwips === null) {
        const { trHeight: _h, ...rest } = props;
        return rest;
      }
      const value = Math.max(0, Math.floor(payload.heightTwips));
      return {
        ...props,
        trHeight: { value, ...(payload.rule ? { rule: payload.rule } : {}) },
      };
    }, "row-height");
  },
};

export interface SetColumnWidthPayload {
  readonly tableId: string;
  /** 0-based grid column index. */
  readonly column: number;
  /** Column width in twips. */
  readonly widthTwips: number;
}

export const setColumnWidthHandler: CommandHandler<SetColumnWidthPayload, DocxSnapshot> = {
  type: "docx:set-column-width",
  apply(snapshot, payload) {
    if (!Number.isFinite(payload.widthTwips) || payload.widthTwips <= 0) {
      throw new CommandError("invalid-payload", `widthTwips must be a positive number, got ${payload.widthTwips}`);
    }
    const located = locateTable(snapshot.root, payload.tableId);
    const table = located.table;
    if (payload.column < 0 || payload.column >= table.grid.length) {
      throw new CommandError(
        "invalid-position",
        `column index ${payload.column} out of range [0, ${table.grid.length}) for table "${payload.tableId}"`
      );
    }

    const w = Math.floor(payload.widthTwips);
    const grid = table.grid.slice();
    if (grid[payload.column]?.w === w) {
      return {
        next: snapshot,
        diff: buildDiff(snapshot.revision, snapshot.revision, {
          kind: "node-updated",
          nodeId: table.id,
          path: ["body", located.bodyIndex, "grid", payload.column],
          field: "column-width",
          summary: "noop",
        }),
      };
    }
    grid[payload.column] = { w };

    // Mirror onto the corresponding tcW for cells that span exactly one
    // grid column at this index. Spanned cells keep their explicit tcW.
    const newRows = table.rows.map((row): TableRow => {
      const cells = row.cells.slice();
      let cursor = 0;
      let mutated = false;
      for (let i = 0; i < cells.length; i++) {
        const cell = cells[i];
        const span = cell.properties.gridSpan ?? 1;
        if (cursor === payload.column && span === 1) {
          cells[i] = {
            ...cell,
            properties: {
              ...cell.properties,
              tcW: { value: w, type: "dxa" },
            },
          };
          mutated = true;
        }
        cursor += span;
      }
      return mutated ? { ...row, cells } : row;
    });

    const next = withoutRaw({ ...table, grid, rows: newRows });
    return commit(snapshot, located, next, payload.tableId, "column-width");
  },
};

export interface MergeCellsHorizontalPayload {
  readonly tableId: string;
  readonly row: number;
  /** 0-based starting grid column. */
  readonly fromColumn: number;
  /** 0-based ending grid column (inclusive). */
  readonly toColumn: number;
}

export const mergeCellsHorizontalHandler: CommandHandler<MergeCellsHorizontalPayload, DocxSnapshot> = {
  type: "docx:merge-cells-horizontal",
  apply(snapshot, payload) {
    if (payload.toColumn < payload.fromColumn) {
      throw new CommandError("invalid-payload", "toColumn must be >= fromColumn");
    }
    const located = locateTable(snapshot.root, payload.tableId);
    const table = located.table;
    const row = table.rows[payload.row];
    if (!row) {
      throw new CommandError("invalid-position", `row ${payload.row} out of range for table "${payload.tableId}"`);
    }

    const span = payload.toColumn - payload.fromColumn + 1;
    const cells = row.cells.slice();

    // Walk the row to find the cells covering [fromColumn, toColumn].
    let cursor = 0;
    let firstIdx = -1;
    let removeFromIdx = -1;
    let removeCount = 0;
    for (let i = 0; i < cells.length; i++) {
      const cellSpan = cells[i].properties.gridSpan ?? 1;
      const startCol = cursor;
      const endCol = cursor + cellSpan - 1;
      if (firstIdx === -1 && startCol <= payload.fromColumn && endCol >= payload.fromColumn) {
        firstIdx = i;
        removeFromIdx = i + 1;
      }
      if (firstIdx !== -1 && i > firstIdx && startCol <= payload.toColumn) {
        removeCount += 1;
      }
      cursor += cellSpan;
    }
    if (firstIdx === -1) {
      throw new CommandError(
        "invalid-position",
        `column ${payload.fromColumn} out of range for row ${payload.row} of table "${payload.tableId}"`
      );
    }

    const merged: TableCell = {
      ...cells[firstIdx],
      properties: { ...cells[firstIdx].properties, gridSpan: span },
    };
    cells.splice(removeFromIdx, removeCount);
    cells[firstIdx] = merged;

    const newRow: TableRow = { ...row, cells };
    const newRows = table.rows.slice();
    newRows[payload.row] = newRow;
    const next = withoutRaw({ ...table, rows: newRows });
    return commit(snapshot, located, next, payload.tableId, "merge-horizontal");
  },
};

// ─── helpers ──────────────────────────────────────────────────────────

function locateTable(root: DocxDocument, tableId: string): TableLocation {
  if (!tableId) {
    throw new CommandError("unknown-target", "tableId must be a non-empty string");
  }
  const located = findTable(root, tableId);
  if (!located) {
    throw new CommandError("unknown-target", `no table with id "${tableId}"`);
  }
  return located;
}

function mutateCell(
  snapshot: DocxSnapshot,
  payload: { tableId: string; row: number; column: number },
  patch: (props: TableCellProperties) => TableCellProperties,
  field: string
) {
  const located = locateTable(snapshot.root, payload.tableId);
  const table = located.table;
  const row = table.rows[payload.row];
  if (!row) {
    throw new CommandError(
      "invalid-position",
      `row ${payload.row} out of range for table "${payload.tableId}"`
    );
  }

  // Resolve column index → cell index (skipping over horizontal-merge spans).
  let cursor = 0;
  let cellIdx = -1;
  for (let i = 0; i < row.cells.length; i++) {
    const span = row.cells[i].properties.gridSpan ?? 1;
    if (cursor <= payload.column && payload.column < cursor + span) {
      cellIdx = i;
      break;
    }
    cursor += span;
  }
  if (cellIdx === -1) {
    throw new CommandError(
      "invalid-position",
      `column ${payload.column} out of range for row ${payload.row} of table "${payload.tableId}"`
    );
  }

  const cell = row.cells[cellIdx];
  const nextProps = patch(cell.properties);
  if (nextProps === cell.properties) {
    return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, { kind: "node-updated", nodeId: cell.id, path: [], field, summary: "noop" }) };
  }

  const nextCell: TableCell = { ...cell, properties: nextProps };
  const nextCells = row.cells.slice();
  nextCells[cellIdx] = nextCell;
  const nextRow: TableRow = { ...row, cells: nextCells };
  const nextRows = table.rows.slice();
  nextRows[payload.row] = nextRow;
  const next = withoutRaw({ ...table, rows: nextRows });
  return commit(snapshot, located, next, payload.tableId, field);
}

function mutateRow(
  snapshot: DocxSnapshot,
  payload: { tableId: string; row: number },
  patch: (props: TableRowProperties) => TableRowProperties,
  field: string
) {
  const located = locateTable(snapshot.root, payload.tableId);
  const table = located.table;
  const row = table.rows[payload.row];
  if (!row) {
    throw new CommandError(
      "invalid-position",
      `row ${payload.row} out of range for table "${payload.tableId}"`
    );
  }
  const nextProps = patch(row.properties);
  const nextRow: TableRow = { ...row, properties: nextProps };
  const nextRows = table.rows.slice();
  nextRows[payload.row] = nextRow;
  const next = withoutRaw({ ...table, rows: nextRows });
  return commit(snapshot, located, next, payload.tableId, field);
}

function commit(
  snapshot: DocxSnapshot,
  located: TableLocation,
  nextTable: Table,
  tableId: string,
  field: string
) {
  const nextDoc = located.replace(nextTable);
  const next = evolveSnapshot(snapshot, nextDoc, { body: true });
  return {
    next,
    diff: buildDiff(snapshot.revision, next.revision, {
      kind: "node-updated",
      nodeId: nextTable.id,
      path: ["body", located.bodyIndex, "table"],
      field,
      summary: `${field} on table "${tableId}"`,
    }),
  };
}

function normaliseHex(value: string): string {
  const v = value.trim().replace(/^#/, "").toUpperCase();
  if (!/^[0-9A-F]{6}$/.test(v) && v !== "AUTO") {
    throw new CommandError("invalid-payload", `Expected hex RGB color (e.g. "FF8800") or "auto", got "${value}"`);
  }
  return v;
}
