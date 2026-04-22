import { ooxml } from "@officeai/core";
import {
  PIVOT_CACHE_DEFINITION_CONTENT_TYPE,
  PIVOT_CACHE_RECORDS_CONTENT_TYPE,
  PIVOT_TABLE_CONTENT_TYPE,
} from "../parser/pivot-tables.js";
import type { PivotCachePart, PivotTablePart, XlsxWorkbook } from "../model/types.js";

/**
 * Phase 1 pivot-table emitter. We re-emit the typed pivot parts
 * byte-identical from `raw` whenever `dirty.pivotTables` is `false`
 * (always in Phase 1). The container clone we serialize from
 * already carries the original bytes, but we explicitly write back
 * here so that:
 *
 *  1. The wiring is in place for Phase 3 typed edits — flipping
 *     `dirty.pivotTables` to `true` becomes the only knob needed
 *     to swap to a deterministic re-render path; the call sites
 *     don't have to grow new branches.
 *  2. A consumer that reaches in and mutates `pivotTables[i].raw`
 *     before serialize (escape hatch for tooling that wants to
 *     hand-edit a pivot before round-trip) is honoured.
 *  3. `[Content_Types].xml` overrides are reconciled against the
 *     live pivot list so a brand-new pivot part still parses in
 *     Excel; orphan overrides for pivot parts that disappeared
 *     since load are dropped.
 *
 * No-op when the workbook has no pivot tables AND no pivot caches.
 */
export function serializePivotParts(workbook: XlsxWorkbook, container: ooxml.OoxmlContainer): void {
  const hasPivots = workbook.pivotTables.length > 0 || workbook.pivotCaches.length > 0;
  if (!hasPivots) return;

  for (const cache of workbook.pivotCaches) {
    writePivotCachePart(cache, container);
  }
  for (const table of workbook.pivotTables) {
    writePivotTablePart(table, container);
  }

  reconcileContentTypes(workbook, container);
}

function writePivotCachePart(cache: PivotCachePart, container: ooxml.OoxmlContainer): void {
  container.writeText(cache.partPath, cache.raw);

  const relsPath = ooxml.RelationshipGraph.relsPathFor(cache.partPath);
  if (cache.relsXml !== undefined) {
    container.writeText(relsPath, cache.relsXml);
  }

  if (cache.recordsPartPath && cache.recordsRaw !== undefined) {
    container.writeText(cache.recordsPartPath, cache.recordsRaw);
  }
}

function writePivotTablePart(table: PivotTablePart, container: ooxml.OoxmlContainer): void {
  container.writeText(table.partPath, table.raw);

  if (table.relsXml !== undefined) {
    const relsPath = ooxml.RelationshipGraph.relsPathFor(table.partPath);
    container.writeText(relsPath, table.relsXml);
  }
}

/**
 * Make sure every live pivot part has a matching `<Override>` entry
 * in `[Content_Types].xml`, and drop overrides for pivot parts no
 * longer in the workbook. Phase 1 never adds new pivots, so the
 * "add" branches mostly fire for files that were saved without
 * the standard overrides (rare, but seen in older third-party
 * authoring tools).
 */
function reconcileContentTypes(workbook: XlsxWorkbook, container: ooxml.OoxmlContainer): void {
  const ct = ooxml.ContentTypes.load(container);
  let mutated = false;

  const livePivotTablePartNames = new Set(workbook.pivotTables.map((p) => `/${p.partPath}`));
  const livePivotCachePartNames = new Set(workbook.pivotCaches.map((p) => `/${p.partPath}`));
  const livePivotRecordsPartNames = new Set(
    workbook.pivotCaches.filter((p) => p.recordsPartPath).map((p) => `/${p.recordsPartPath}`)
  );

  for (const partName of livePivotTablePartNames) {
    if (!ct.hasOverride(partName)) {
      ct.addOverride(partName, PIVOT_TABLE_CONTENT_TYPE);
      mutated = true;
    }
  }
  for (const partName of livePivotCachePartNames) {
    if (!ct.hasOverride(partName)) {
      ct.addOverride(partName, PIVOT_CACHE_DEFINITION_CONTENT_TYPE);
      mutated = true;
    }
  }
  for (const partName of livePivotRecordsPartNames) {
    if (!ct.hasOverride(partName)) {
      ct.addOverride(partName, PIVOT_CACHE_RECORDS_CONTENT_TYPE);
      mutated = true;
    }
  }

  for (const o of [...ct.overrides]) {
    if (o.contentType === PIVOT_TABLE_CONTENT_TYPE && !livePivotTablePartNames.has(o.partName)) {
      ct.removeOverride(o.partName);
      mutated = true;
    } else if (
      o.contentType === PIVOT_CACHE_DEFINITION_CONTENT_TYPE &&
      !livePivotCachePartNames.has(o.partName)
    ) {
      ct.removeOverride(o.partName);
      mutated = true;
    } else if (
      o.contentType === PIVOT_CACHE_RECORDS_CONTENT_TYPE &&
      !livePivotRecordsPartNames.has(o.partName)
    ) {
      ct.removeOverride(o.partName);
      mutated = true;
    }
  }

  if (mutated) ct.writeBack(container);
}
