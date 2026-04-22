import { ooxml } from "@officeai/core";
import type { PivotCachePart, PivotTablePart, Sheet } from "../model/types.js";
import { resolveTargetPath } from "./parse.js";

/**
 * Phase 1 pivot-table preservation, per
 * `spec/xlsx/pivot-tables.md`. We promote `xl/pivotTables/*` and
 * `xl/pivotCache/*` from the catch-all `opaqueParts` bucket to two
 * dedicated typed slots on the workbook so consumers can identify
 * pivots by name + cacheId without re-walking the OPC graph. The
 * parts themselves stay opaque — the entire `<pivotTableDefinition>`
 * / `<pivotCacheDefinition>` element rides verbatim on
 * {@link PivotTablePart.raw} / {@link PivotCachePart.raw} so the
 * serializer can re-emit them byte-identical when no typed edits
 * have occurred (always the case in Phase 1).
 *
 * Phase 3 will replace the `raw` blobs with full typed fields
 * (rows / cols / data / filter axes, cache field schemas,
 * materialised records). Until then, the only typed metadata we
 * lift is what the serializer needs for round-trip discipline plus
 * what consumers need to render a pivot picker:
 *
 *   - PivotTablePart: `name` (display), `cacheId` (link to cache)
 *   - PivotCachePart: `cacheId`, optional `recordsPartPath`
 */

const REL_PIVOT_CACHE_DEFINITION =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
const REL_PIVOT_CACHE_RECORDS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
const REL_PIVOT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";

const WORKBOOK_PART = "xl/workbook.xml";

export const PIVOT_TABLE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
export const PIVOT_CACHE_DEFINITION_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
export const PIVOT_CACHE_RECORDS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";

export interface DiscoveredPivotParts {
  readonly pivotTables: ReadonlyArray<PivotTablePart>;
  readonly pivotCaches: ReadonlyArray<PivotCachePart>;
  /**
   * Every container path the typed pivot model now owns. The parser
   * subtracts this from `opaqueParts` so the bytes aren't double-
   * tracked, and the serializer re-emits them via
   * `serializePivotParts`. Includes definition parts, records parts,
   * and any matching `_rels/*.rels` sidecars.
   */
  readonly modeledPaths: ReadonlySet<string>;
}

/**
 * Walk the workbook + sheet rels graph for every pivot-related part
 * and return them as typed records. We never hardcode part paths —
 * Excel and LibreOffice don't always agree on the `xl/pivotCache/`
 * directory casing, and external authoring tools occasionally place
 * the cache definition somewhere unusual. The rels graph is the
 * source of truth.
 */
export function discoverPivotParts(
  container: ooxml.OoxmlContainer,
  contentTypes: ReadonlyMap<string, string>,
  sheets: ReadonlyArray<Sheet>
): DiscoveredPivotParts {
  const modeledPaths = new Set<string>();

  const pivotCaches = discoverPivotCaches(container, contentTypes, modeledPaths);
  const pivotTables = discoverPivotTables(container, contentTypes, sheets, modeledPaths);

  return { pivotTables, pivotCaches, modeledPaths };
}

function discoverPivotCaches(
  container: ooxml.OoxmlContainer,
  contentTypes: ReadonlyMap<string, string>,
  modeledPaths: Set<string>
): ReadonlyArray<PivotCachePart> {
  const out: PivotCachePart[] = [];
  if (!container.has(WORKBOOK_PART)) return out;

  const workbookRels = ooxml.RelationshipGraph.loadFor(container, WORKBOOK_PART);
  const cacheRels = workbookRels.relationships.filter((r) => r.type === REL_PIVOT_CACHE_DEFINITION);
  if (cacheRels.length === 0) return out;

  // Track the cacheId discovered from `<workbook><pivotCaches>` so a
  // cache part that omits its own `cacheId` attribute still gets one
  // matching what the workbook sees (Excel always writes the
  // workbook-side `cacheId`; the cache-part-side attribute is
  // optional in the OOXML schema).
  const cacheIdByRid = readWorkbookPivotCacheIds(container);
  // Fallback ids when neither the workbook nor the cache part
  // carries one — assign sequentially in rels order so the result
  // is deterministic.
  let fallbackId = 0;

  for (const rel of cacheRels) {
    const partPath = resolveTargetPath(WORKBOOK_PART, rel.target);
    if (!container.has(partPath)) continue;

    const raw = container.readText(partPath);
    const cacheIdFromXml = extractRootIntAttribute(raw, "pivotCacheDefinition", "cacheId");
    const cacheIdFromWorkbook = cacheIdByRid.get(rel.id);
    const cacheId = cacheIdFromXml ?? cacheIdFromWorkbook ?? fallbackId;
    if (cacheIdFromXml === undefined && cacheIdFromWorkbook === undefined) {
      fallbackId++;
    }

    modeledPaths.add(partPath);

    const relsPath = ooxml.RelationshipGraph.relsPathFor(partPath);
    const relsXml = container.has(relsPath) ? container.readText(relsPath) : undefined;
    if (container.has(relsPath)) modeledPaths.add(relsPath);

    let recordsPartPath: string | undefined;
    let recordsRaw: string | undefined;
    let recordsContentType: string | undefined;
    if (container.has(relsPath)) {
      const cacheRelsGraph = ooxml.RelationshipGraph.loadFor(container, partPath);
      const recordsRel = cacheRelsGraph.relationships.find((r) => r.type === REL_PIVOT_CACHE_RECORDS);
      if (recordsRel) {
        const recPath = resolveTargetPath(partPath, recordsRel.target);
        if (container.has(recPath)) {
          recordsPartPath = recPath;
          recordsRaw = container.readText(recPath);
          recordsContentType = contentTypes.get(recPath);
          modeledPaths.add(recPath);
          const recordsRelsPath = ooxml.RelationshipGraph.relsPathFor(recPath);
          if (container.has(recordsRelsPath)) modeledPaths.add(recordsRelsPath);
        }
      }
    }

    out.push({
      partPath,
      cacheId,
      raw,
      ...(contentTypes.get(partPath) ? { contentType: contentTypes.get(partPath)! } : {}),
      ...(relsXml !== undefined ? { relsXml } : {}),
      ...(recordsPartPath ? { recordsPartPath } : {}),
      ...(recordsRaw !== undefined ? { recordsRaw } : {}),
      ...(recordsContentType ? { recordsContentType } : {}),
    });
  }
  return out;
}

function discoverPivotTables(
  container: ooxml.OoxmlContainer,
  contentTypes: ReadonlyMap<string, string>,
  sheets: ReadonlyArray<Sheet>,
  modeledPaths: Set<string>
): ReadonlyArray<PivotTablePart> {
  const out: PivotTablePart[] = [];
  const seenPartPaths = new Set<string>();

  for (const sheet of sheets) {
    if (sheet.kind !== "worksheet") continue;
    const sheetRelsPath = ooxml.RelationshipGraph.relsPathFor(sheet.partPath);
    if (!container.has(sheetRelsPath)) continue;

    const sheetRels = ooxml.RelationshipGraph.loadFor(container, sheet.partPath);
    const tableRels = sheetRels.relationships.filter((r) => r.type === REL_PIVOT_TABLE);
    if (tableRels.length === 0) continue;

    for (const rel of tableRels) {
      const partPath = resolveTargetPath(sheet.partPath, rel.target);
      if (!container.has(partPath)) continue;
      if (seenPartPaths.has(partPath)) continue;
      seenPartPaths.add(partPath);

      const raw = container.readText(partPath);
      const name = extractRootStringAttribute(raw, "pivotTableDefinition", "name") ?? partPath;
      const cacheId = extractRootIntAttribute(raw, "pivotTableDefinition", "cacheId");

      modeledPaths.add(partPath);
      const relsPath = ooxml.RelationshipGraph.relsPathFor(partPath);
      const relsXml = container.has(relsPath) ? container.readText(relsPath) : undefined;
      if (container.has(relsPath)) modeledPaths.add(relsPath);

      out.push({
        partPath,
        name,
        ...(cacheId !== undefined ? { cacheId } : {}),
        raw,
        ...(contentTypes.get(partPath) ? { contentType: contentTypes.get(partPath)! } : {}),
        ...(relsXml !== undefined ? { relsXml } : {}),
      });
    }
  }
  return out;
}

/**
 * Pull `<pivotCaches><pivotCache cacheId="N" r:id="rIdM"/></pivotCaches>`
 * out of `xl/workbook.xml`. We use this to attach the workbook-side
 * `cacheId` to a cache part whose own root element omits it.
 */
function readWorkbookPivotCacheIds(container: ooxml.OoxmlContainer): Map<string, number> {
  const out = new Map<string, number>();
  if (!container.has(WORKBOOK_PART)) return out;
  const xml = container.readText(WORKBOOK_PART);
  const block = /<pivotCaches\b[^>]*>([\s\S]*?)<\/pivotCaches>/.exec(xml);
  if (!block) return out;
  const inner = block[1] ?? "";
  const re = /<pivotCache\b([^/>]*)\/?>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(inner)) !== null) {
    const attrs = m[1] ?? "";
    const idMatch = /\bcacheId=("|')([^"']+)\1/.exec(attrs);
    const ridMatch = /\br:id=("|')([^"']+)\1/.exec(attrs) ?? /\br:Id=("|')([^"']+)\1/.exec(attrs);
    if (!idMatch || !ridMatch) continue;
    const cid = Number.parseInt(idMatch[2]!, 10);
    if (!Number.isFinite(cid)) continue;
    out.set(ridMatch[2]!, cid);
  }
  return out;
}

/**
 * Pull a numeric attribute off the root element. Returns `undefined`
 * when the attribute is missing or non-numeric — the caller decides
 * whether that's an error or a recoverable defect.
 *
 * We use a regex rather than the full XML parser because Phase 1
 * doesn't need to walk the body and the root element of a pivot
 * part is always a single element with attributes on a couple of
 * known names; a regex is sufficient and avoids paying the parse
 * cost for a part we're going to round-trip verbatim anyway.
 */
function extractRootIntAttribute(xml: string, tag: string, attr: string): number | undefined {
  const v = extractRootStringAttribute(xml, tag, attr);
  if (v === undefined) return undefined;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n : undefined;
}

function extractRootStringAttribute(xml: string, tag: string, attr: string): string | undefined {
  const tagRe = new RegExp(`<${tag}\\b([^>]*)>`);
  const m = tagRe.exec(xml);
  if (!m) return undefined;
  const attrs = m[1] ?? "";
  const aRe = new RegExp(`\\b${attr}=("|')([^"']*)\\1`);
  const am = aRe.exec(attrs);
  return am ? am[2] : undefined;
}
