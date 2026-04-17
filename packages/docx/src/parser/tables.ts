import { ooxml, type IdMinter } from "@officeai/core";
import type {
  BlockNode,
  Paragraph,
  Table,
  TableCell,
  TableCellProperties,
  TableGridCol,
  TableProperties,
  TableRow,
  TableRowProperties,
  TableWidth,
} from "../model/types.js";
import { attrOf, captureOpaque, elementEntries, findElementEntry } from "./xml-helpers.js";

/**
 * Parser for typed tables (P1.3 / W7).
 *
 * `parseParagraph` is injected to avoid an import cycle with `parse.ts`.
 * Nested tables are handled internally — a `<w:tbl>` encountered inside a
 * `<w:tc>` recurses into `parseTable` directly.
 *
 * Byte-preservation: every `Table` returned by this parser carries the
 * original `<w:tbl>` subtree in `raw`. The serializer keys off `raw` to
 * decide whether to re-emit cached bytes (clean) or regenerate from the
 * typed model (touched). Mutation commands MUST drop `raw` on the new
 * `Table` they produce.
 */

type ParseParagraph = (entry: Record<string, unknown>, mintNodeId: IdMinter) => Paragraph;

export function parseTable(
  entry: Record<string, unknown>,
  mintNodeId: IdMinter,
  parseParagraph: ParseParagraph
): Table {
  const children = (entry["w:tbl"] as unknown[] | undefined) ?? [];
  const tblPr = findElementEntry(children, "w:tblPr");
  const tblGrid = findElementEntry(children, "w:tblGrid");

  const properties = tblPr ? parseTableProperties(tblPr) : {};
  const grid = tblGrid ? parseTableGrid(tblGrid) : [];
  const rows: TableRow[] = [];
  for (const c of elementEntries(children)) {
    if (ooxml.getTag(c) !== "w:tr") continue;
    rows.push(parseTableRow(c, mintNodeId, parseParagraph));
  }

  return {
    kind: "table",
    id: mintNodeId(),
    properties,
    grid,
    rows,
    raw: captureOpaque(entry),
  };
}

function parseTableProperties(entry: Record<string, unknown>): TableProperties {
  const children = (entry["w:tblPr"] as unknown[] | undefined) ?? [];
  const props: { -readonly [K in keyof TableProperties]: TableProperties[K] } = {};
  const opaqueProps: ReturnType<typeof captureOpaque>[] = [];
  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    switch (tag) {
      case "w:tblW": {
        const w = parseWidth(c);
        if (w) props.width = w;
        else opaqueProps.push(captureOpaque(c));
        break;
      }
      case "w:jc": {
        const v = attrOf(c, "w:val");
        if (v === "left" || v === "center" || v === "right" || v === "start" || v === "end") {
          props.jc = v;
        } else {
          opaqueProps.push(captureOpaque(c));
        }
        break;
      }
      default:
        opaqueProps.push(captureOpaque(c));
        break;
    }
  }
  if (opaqueProps.length > 0) props.opaqueProps = opaqueProps;
  return props;
}

function parseTableGrid(entry: Record<string, unknown>): TableGridCol[] {
  const out: TableGridCol[] = [];
  const children = (entry["w:tblGrid"] as unknown[] | undefined) ?? [];
  for (const c of elementEntries(children)) {
    if (ooxml.getTag(c) !== "w:gridCol") continue;
    const wAttr = attrOf(c, "w:w");
    out.push(wAttr !== undefined ? { w: Number(wAttr) } : {});
  }
  return out;
}

export function parseTableRow(
  entry: Record<string, unknown>,
  mintNodeId: IdMinter,
  parseParagraph: ParseParagraph
): TableRow {
  const children = (entry["w:tr"] as unknown[] | undefined) ?? [];
  const trPr = findElementEntry(children, "w:trPr");
  const properties = trPr ? parseTableRowProperties(trPr) : {};
  const cells: TableCell[] = [];
  for (const c of elementEntries(children)) {
    if (ooxml.getTag(c) !== "w:tc") continue;
    cells.push(parseTableCell(c, mintNodeId, parseParagraph));
  }
  return {
    kind: "table-row",
    id: mintNodeId(),
    properties,
    cells,
  };
}

function parseTableRowProperties(entry: Record<string, unknown>): TableRowProperties {
  const children = (entry["w:trPr"] as unknown[] | undefined) ?? [];
  const props: { -readonly [K in keyof TableRowProperties]: TableRowProperties[K] } = {};
  const opaqueProps: ReturnType<typeof captureOpaque>[] = [];
  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    switch (tag) {
      case "w:trHeight": {
        const v = attrOf(c, "w:val");
        const rule = attrOf(c, "w:hRule");
        if (v !== undefined) {
          const trHeight: { value: number; rule?: "auto" | "exact" | "atLeast" } = {
            value: Number(v),
          };
          if (rule === "auto" || rule === "exact" || rule === "atLeast") trHeight.rule = rule;
          props.trHeight = trHeight;
        } else {
          opaqueProps.push(captureOpaque(c));
        }
        break;
      }
      case "w:tblHeader": {
        // `<w:tblHeader/>` (or `<w:tblHeader w:val="true"/>`) → header row.
        // `<w:tblHeader w:val="false"/>` is "explicitly not a header" — we
        // model both presence and explicit-false to preserve round-trip.
        const v = attrOf(c, "w:val");
        props.header = v === "false" || v === "0" || v === "off" ? false : true;
        break;
      }
      default:
        opaqueProps.push(captureOpaque(c));
        break;
    }
  }
  if (opaqueProps.length > 0) props.opaqueProps = opaqueProps;
  return props;
}

export function parseTableCell(
  entry: Record<string, unknown>,
  mintNodeId: IdMinter,
  parseParagraph: ParseParagraph
): TableCell {
  const children = (entry["w:tc"] as unknown[] | undefined) ?? [];
  const tcPr = findElementEntry(children, "w:tcPr");
  const properties = tcPr ? parseTableCellProperties(tcPr) : {};
  const body: BlockNode[] = [];
  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    if (tag === "w:tcPr") continue;
    if (tag === "w:p") {
      body.push(parseParagraph(c, mintNodeId));
    } else if (tag === "w:tbl") {
      body.push(parseTable(c, mintNodeId, parseParagraph));
    } else {
      // Unknown block-level child inside a cell — preserve verbatim.
      body.push({ kind: "opaque-block", id: mintNodeId(), raw: captureOpaque(c) });
    }
  }
  // OOXML requires every `<w:tc>` to contain at least one paragraph; if a
  // pathological cell was empty we fabricate one so downstream code can
  // assume `cell.body[0]` exists.
  if (body.length === 0) {
    body.push({
      kind: "paragraph",
      id: mintNodeId(),
      properties: {},
      children: [{ kind: "run", id: mintNodeId(), properties: {}, children: [] }],
    });
  }
  return {
    kind: "table-cell",
    id: mintNodeId(),
    properties,
    body,
  };
}

function parseTableCellProperties(entry: Record<string, unknown>): TableCellProperties {
  const children = (entry["w:tcPr"] as unknown[] | undefined) ?? [];
  const props: { -readonly [K in keyof TableCellProperties]: TableCellProperties[K] } = {};
  const opaqueProps: ReturnType<typeof captureOpaque>[] = [];
  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    switch (tag) {
      case "w:gridSpan": {
        const v = attrOf(c, "w:val");
        const n = v !== undefined ? Number(v) : NaN;
        if (Number.isFinite(n) && n > 0) props.gridSpan = n;
        else opaqueProps.push(captureOpaque(c));
        break;
      }
      case "w:vMerge": {
        const v = attrOf(c, "w:val");
        // OOXML: missing w:val on w:vMerge means "continue".
        if (v === "restart") props.vMerge = "restart";
        else props.vMerge = "continue";
        break;
      }
      case "w:tcW": {
        const w = parseWidth(c);
        if (w) props.tcW = w;
        else opaqueProps.push(captureOpaque(c));
        break;
      }
      default:
        opaqueProps.push(captureOpaque(c));
        break;
    }
  }
  if (opaqueProps.length > 0) props.opaqueProps = opaqueProps;
  return props;
}

function parseWidth(entry: Record<string, unknown>): TableWidth | null {
  const valueAttr = attrOf(entry, "w:w");
  const typeAttr = attrOf(entry, "w:type") ?? "dxa";
  if (valueAttr === undefined) return null;
  // The `w:w` attribute on `<w:tblW>` with `type="pct"` can be either a raw
  // 50ths-of-a-percent integer or a `"100%"` literal. Strip a trailing `%`
  // and treat it as the raw value if present (Word writes both forms).
  const raw = valueAttr.endsWith("%") ? valueAttr.slice(0, -1) : valueAttr;
  const n = Number(raw);
  if (!Number.isFinite(n)) return null;
  const type =
    typeAttr === "auto" || typeAttr === "dxa" || typeAttr === "pct" || typeAttr === "nil" ? typeAttr : "dxa";
  return { value: n, type };
}
