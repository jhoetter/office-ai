import type {
  HeaderFooterRef,
  PageColumns,
  PageMargins,
  PageSize,
  SectionProperties,
} from "../model/types.js";
import { opaqueToEntry } from "../parser/xml-helpers.js";

/**
 * Build a `<w:sectPr>` entry from a typed {@link SectionProperties}.
 *
 * Used when {@link SectionBreak.raw} has been dropped (mutating command
 * produced a fresh node). When `raw` is still present the caller emits
 * it verbatim instead — that is what preserves byte-identical
 * round-trip for sections that were read but never modified.
 */
export function serializeSectionProperties(props: SectionProperties): Record<string, unknown> {
  const children: unknown[] = [];

  for (const ref of props.headerRefs) {
    children.push(serializeHeaderFooterRef("w:headerReference", ref));
  }
  for (const ref of props.footerRefs) {
    children.push(serializeHeaderFooterRef("w:footerReference", ref));
  }

  if (props.titlePg) children.push({ "w:titlePg": [] });
  if (props.sectionType) {
    children.push(makeEntry("w:type", [], { "w:val": props.sectionType }));
  }
  if (props.pgSz) children.push(serializePageSize(props.pgSz));
  if (props.pgMar) children.push(serializePageMargins(props.pgMar));
  if (props.cols) children.push(serializePageColumns(props.cols));

  if (props.opaqueProps) {
    for (const op of props.opaqueProps) {
      children.push(opaqueToEntry(op));
    }
  }

  return { "w:sectPr": children };
}

function serializePageSize(pg: PageSize): unknown {
  const attrs: Record<string, string> = {
    "w:w": String(pg.w),
    "w:h": String(pg.h),
  };
  if (pg.orient) attrs["w:orient"] = pg.orient;
  return makeEntry("w:pgSz", [], attrs);
}

function serializePageMargins(m: PageMargins): unknown {
  const attrs: Record<string, string> = {
    "w:top": String(m.top),
    "w:right": String(m.right),
    "w:bottom": String(m.bottom),
    "w:left": String(m.left),
    "w:header": String(m.header),
    "w:footer": String(m.footer),
  };
  if (m.gutter !== undefined) attrs["w:gutter"] = String(m.gutter);
  return makeEntry("w:pgMar", [], attrs);
}

function serializePageColumns(c: PageColumns): unknown {
  const attrs: Record<string, string> = { "w:num": String(c.num) };
  if (c.sep !== undefined) attrs["w:sep"] = c.sep ? "1" : "0";
  if (c.equalWidth !== undefined) attrs["w:equalWidth"] = c.equalWidth ? "1" : "0";
  if (c.space !== undefined) attrs["w:space"] = String(c.space);
  return makeEntry("w:cols", [], attrs);
}

function serializeHeaderFooterRef(tag: string, ref: HeaderFooterRef): unknown {
  return makeEntry(tag, [], { "w:type": ref.type, "r:id": ref.relationshipId });
}

function makeEntry(tag: string, children: unknown[], attrs: Record<string, string>): unknown {
  const entry: Record<string, unknown> = { [tag]: children };
  if (Object.keys(attrs).length > 0) {
    const out: Record<string, string> = {};
    for (const [k, v] of Object.entries(attrs)) {
      out[`@_${k}`] = v;
    }
    entry[":@"] = out;
  }
  return entry;
}
