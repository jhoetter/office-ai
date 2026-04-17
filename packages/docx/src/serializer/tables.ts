import type {
  BlockNode,
  OpaqueXml,
  Table,
  TableCell,
  TableCellProperties,
  TableGridCol,
  TableProperties,
  TableRow,
  TableRowProperties,
  TableWidth,
} from "../model/types.js";
import { opaqueToEntry } from "../parser/xml-helpers.js";

/**
 * Serializer for typed tables (P1.3 / W7).
 *
 * Byte-preservation contract: callers (the body serializer) check
 * `table.raw` BEFORE calling this function. If `raw` is set, they re-emit
 * the cached subtree via `opaqueToEntry` directly. This function only runs
 * for tables whose typed model has been touched and must be regenerated.
 *
 * `serializeBlock` is injected (rather than imported) so we can share
 * paragraph emission with the main serializer without an import cycle. It
 * already performs the raw-blob shortcut for nested tables.
 */
export function serializeTable(table: Table, serializeBlock: (block: BlockNode) => unknown): unknown {
  const children: unknown[] = [];
  const tblPr = serializeTableProperties(table.properties);
  if (tblPr) children.push(tblPr);
  if (table.grid.length > 0) {
    children.push(serializeTableGrid(table.grid));
  }
  for (const row of table.rows) {
    children.push(serializeTableRow(row, serializeBlock));
  }
  return { "w:tbl": children };
}

function serializeTableProperties(props: TableProperties): unknown | null {
  const out: unknown[] = [];
  if (props.width) {
    out.push(makeWidthEl("w:tblW", props.width));
  }
  if (props.jc) {
    out.push(makeEl("w:jc", { "w:val": props.jc }));
  }
  if (props.opaqueProps) {
    for (const o of props.opaqueProps) {
      out.push(opaqueToEntry(o));
    }
  }
  if (out.length === 0) return null;
  return { "w:tblPr": out };
}

function serializeTableGrid(grid: ReadonlyArray<TableGridCol>): unknown {
  const cols: unknown[] = grid.map((g) =>
    g.w !== undefined ? makeEl("w:gridCol", { "w:w": String(g.w) }) : { "w:gridCol": [] }
  );
  return { "w:tblGrid": cols };
}

function serializeTableRow(row: TableRow, serializeBlock: (b: BlockNode) => unknown): unknown {
  const children: unknown[] = [];
  const trPr = serializeTableRowProperties(row.properties);
  if (trPr) children.push(trPr);
  for (const cell of row.cells) {
    children.push(serializeTableCell(cell, serializeBlock));
  }
  return { "w:tr": children };
}

function serializeTableRowProperties(props: TableRowProperties): unknown | null {
  const out: unknown[] = [];
  if (props.trHeight) {
    const attrs: Record<string, string> = { "w:val": String(props.trHeight.value) };
    if (props.trHeight.rule) attrs["w:hRule"] = props.trHeight.rule;
    out.push(makeEl("w:trHeight", attrs));
  }
  if (props.header === true) {
    out.push({ "w:tblHeader": [] });
  } else if (props.header === false) {
    out.push(makeEl("w:tblHeader", { "w:val": "false" }));
  }
  if (props.opaqueProps) {
    for (const o of props.opaqueProps) {
      out.push(opaqueToEntry(o));
    }
  }
  if (out.length === 0) return null;
  return { "w:trPr": out };
}

function serializeTableCell(cell: TableCell, serializeBlock: (b: BlockNode) => unknown): unknown {
  const children: unknown[] = [];
  const tcPr = serializeTableCellProperties(cell.properties);
  if (tcPr) children.push(tcPr);
  for (const block of cell.body) {
    children.push(serializeBlock(block));
  }
  return { "w:tc": children };
}

function serializeTableCellProperties(props: TableCellProperties): unknown | null {
  const out: unknown[] = [];
  if (props.tcW) {
    out.push(makeWidthEl("w:tcW", props.tcW));
  }
  if (props.gridSpan !== undefined && props.gridSpan > 1) {
    out.push(makeEl("w:gridSpan", { "w:val": String(props.gridSpan) }));
  }
  if (props.vMerge) {
    if (props.vMerge === "restart") out.push(makeEl("w:vMerge", { "w:val": "restart" }));
    else out.push({ "w:vMerge": [] });
  }
  if (props.opaqueProps) {
    for (const o of props.opaqueProps) {
      out.push(opaqueToEntry(o));
    }
  }
  if (out.length === 0) return null;
  return { "w:tcPr": out };
}

function makeWidthEl(tag: string, w: TableWidth): unknown {
  return makeEl(tag, { "w:w": String(w.value), "w:type": w.type });
}

function makeEl(tag: string, attrs?: Record<string, string>): unknown {
  const entry: Record<string, unknown> = { [tag]: [] };
  if (attrs && Object.keys(attrs).length > 0) {
    const out: Record<string, string> = {};
    for (const [k, v] of Object.entries(attrs)) {
      out[`@_${k}`] = v;
    }
    entry[":@"] = out;
  }
  return entry;
}

/**
 * Re-emit a table that hasn't been touched since parse using its cached
 * `raw` subtree. Called by the main body serializer as the fast path.
 */
export function serializeTableFromRaw(raw: OpaqueXml): unknown {
  return opaqueToEntry(raw);
}
