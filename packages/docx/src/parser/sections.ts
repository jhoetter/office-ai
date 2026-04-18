import { ooxml } from "@officeai/core";
import type {
  HeaderFooterRef,
  PageColumns,
  PageMargins,
  PageSize,
  SectionProperties,
} from "../model/types.js";
import { attrOf, captureOpaque, elementEntries } from "./xml-helpers.js";

/**
 * Parse a `<w:sectPr>` element entry into a typed
 * {@link SectionProperties} record.
 *
 * Only the fields the renderer reasons about are promoted to typed
 * shape. Everything else (`<w:lineNumType>`, `<w:pgNumType>`,
 * `<w:formProt>`, `<w:vAlign>`, `<w:rtlGutter>`, `<w:docGrid>`, …) is
 * captured into `opaqueProps` in original document order so the rebuild
 * serializer can re-emit it without dropping bytes.
 *
 * Round-trip contract: `SectionBreak.raw` is the source of truth for
 * untouched sections. This typed projection is what mutating commands
 * write through; the serializer falls back to `raw` when nothing has
 * changed.
 */
export function parseSectionProperties(entry: Record<string, unknown>): SectionProperties {
  const tag = ooxml.getTag(entry);
  const children = (entry[tag] as unknown[] | undefined) ?? [];

  let pgSz: PageSize | undefined;
  let pgMar: PageMargins | undefined;
  let cols: PageColumns | undefined;
  let titlePg = false;
  let sectionType: SectionProperties["sectionType"] | undefined;
  const headerRefs: HeaderFooterRef[] = [];
  const footerRefs: HeaderFooterRef[] = [];
  const opaqueProps: ReturnType<typeof captureOpaque>[] = [];

  for (const child of elementEntries(children)) {
    const childTag = ooxml.getTag(child);
    switch (childTag) {
      case "w:pgSz":
        pgSz = parsePageSize(child);
        break;
      case "w:pgMar":
        pgMar = parsePageMargins(child);
        break;
      case "w:cols":
        cols = parsePageColumns(child);
        break;
      case "w:titlePg":
        titlePg = true;
        break;
      case "w:type": {
        const v = attrOf(child, "w:val");
        if (
          v === "continuous" ||
          v === "nextPage" ||
          v === "oddPage" ||
          v === "evenPage" ||
          v === "nextColumn"
        ) {
          sectionType = v;
        } else {
          opaqueProps.push(captureOpaque(child));
        }
        break;
      }
      case "w:headerReference": {
        const ref = parseHeaderFooterRef(child);
        if (ref) headerRefs.push(ref);
        else opaqueProps.push(captureOpaque(child));
        break;
      }
      case "w:footerReference": {
        const ref = parseHeaderFooterRef(child);
        if (ref) footerRefs.push(ref);
        else opaqueProps.push(captureOpaque(child));
        break;
      }
      default:
        opaqueProps.push(captureOpaque(child));
        break;
    }
  }

  const props: SectionProperties = {
    ...(pgSz ? { pgSz } : {}),
    ...(pgMar ? { pgMar } : {}),
    ...(cols ? { cols } : {}),
    headerRefs,
    footerRefs,
    ...(titlePg ? { titlePg } : {}),
    ...(sectionType ? { sectionType } : {}),
    ...(opaqueProps.length > 0 ? { opaqueProps } : {}),
  };
  return props;
}

function parsePageSize(entry: Record<string, unknown>): PageSize | undefined {
  const w = numericAttr(entry, "w:w");
  const h = numericAttr(entry, "w:h");
  if (w === undefined || h === undefined) return undefined;
  const orient = attrOf(entry, "w:orient");
  return {
    w,
    h,
    ...(orient === "portrait" || orient === "landscape" ? { orient } : {}),
  };
}

function parsePageMargins(entry: Record<string, unknown>): PageMargins | undefined {
  const top = numericAttr(entry, "w:top");
  const right = numericAttr(entry, "w:right");
  const bottom = numericAttr(entry, "w:bottom");
  const left = numericAttr(entry, "w:left");
  const header = numericAttr(entry, "w:header");
  const footer = numericAttr(entry, "w:footer");
  if (
    top === undefined ||
    right === undefined ||
    bottom === undefined ||
    left === undefined ||
    header === undefined ||
    footer === undefined
  ) {
    return undefined;
  }
  const gutter = numericAttr(entry, "w:gutter");
  return {
    top,
    right,
    bottom,
    left,
    header,
    footer,
    ...(gutter !== undefined ? { gutter } : {}),
  };
}

function parsePageColumns(entry: Record<string, unknown>): PageColumns | undefined {
  const num = numericAttr(entry, "w:num") ?? 1;
  const sepStr = attrOf(entry, "w:sep");
  const equalWidthStr = attrOf(entry, "w:equalWidth");
  const space = numericAttr(entry, "w:space");
  return {
    num,
    ...(sepStr !== undefined ? { sep: sepStr === "1" || sepStr === "true" } : {}),
    ...(equalWidthStr !== undefined ? { equalWidth: equalWidthStr === "1" || equalWidthStr === "true" } : {}),
    ...(space !== undefined ? { space } : {}),
  };
}

function parseHeaderFooterRef(entry: Record<string, unknown>): HeaderFooterRef | undefined {
  const t = attrOf(entry, "w:type");
  const rId = attrOf(entry, "r:id");
  if (!rId) return undefined;
  const type: HeaderFooterRef["type"] = t === "first" || t === "even" || t === "default" ? t : "default";
  return { type, relationshipId: rId };
}

function numericAttr(entry: Record<string, unknown>, name: string): number | undefined {
  const raw = attrOf(entry, name);
  if (raw === undefined) return undefined;
  const n = Number(raw);
  return Number.isFinite(n) ? n : undefined;
}
