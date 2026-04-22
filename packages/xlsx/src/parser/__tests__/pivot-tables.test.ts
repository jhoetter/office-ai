import { readFile } from "node:fs/promises";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { sha256Hex } from "@officeai/core";
import { parseXlsx } from "../parse.js";
import { serializeXlsx } from "../../serializer/serialize.js";

const here = dirname(fileURLToPath(import.meta.url));
const fixtures = resolve(here, "../../../../../fixtures/xlsx/synthetic");

/**
 * Phase 1 of `spec/xlsx/pivot-tables.md` — preservation only. A pivot
 * table parses into a typed `PivotTablePart` (with `name` + `cacheId`
 * lifted out) and a `PivotCachePart` (with the matching
 * `pivotCacheRecords` blob), and the entire bundle round-trips
 * byte-identical via the serializer when no typed pivot edits land.
 *
 * No fixture in `fixtures/xlsx/` ships a pivot today (the
 * `real-excel-mac-2021-pivot.xlsx` slot in the manifest is reserved
 * for a future Microsoft-emitted file). We synthesize one in-memory
 * by augmenting `01-single-sheet-numbers.xlsx` with a minimal pivot
 * cache + table that points at its `Inventory` sheet.
 */

const PIVOT_TABLE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotInventory" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">
  <location ref="A8:B12" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>
  <pivotFields count="2">
    <pivotField axis="axisRow" showAll="0">
      <items count="3">
        <item x="0"/>
        <item x="1"/>
        <item t="default"/>
      </items>
    </pivotField>
    <pivotField dataField="1" showAll="0"/>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <rowItems count="3">
    <i><x/></i>
    <i><x v="1"/></i>
    <i t="grand"><x/></i>
  </rowItems>
  <colItems count="1"><i/></colItems>
  <dataFields count="1">
    <dataField name="Sum of Qty" fld="1" baseField="0" baseItem="0"/>
  </dataFields>
  <pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>
</pivotTableDefinition>`;

const PIVOT_TABLE_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>`;

const PIVOT_CACHE_DEFINITION_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshedBy="OfficeAI" refreshedDate="46500.0" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="2">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:B3" sheet="Inventory"/>
  </cacheSource>
  <cacheFields count="2">
    <cacheField name="Item" numFmtId="0">
      <sharedItems count="2">
        <s v="Widget"/>
        <s v="Gadget"/>
      </sharedItems>
    </cacheField>
    <cacheField name="Qty" numFmtId="0">
      <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="4" maxValue="10"/>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>`;

const PIVOT_CACHE_DEFINITION_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>`;

const PIVOT_CACHE_RECORDS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="2">
  <r><x v="0"/><n v="10"/></r>
  <r><x v="1"/><n v="4"/></r>
</pivotCacheRecords>`;

async function buildPivotFixture(): Promise<Uint8Array> {
  const baseBytes = await readFile(resolve(fixtures, "01-single-sheet-numbers.xlsx"));
  const zip = await JSZip.loadAsync(baseBytes);

  zip.file("xl/pivotTables/pivotTable1.xml", PIVOT_TABLE_XML);
  zip.file("xl/pivotTables/_rels/pivotTable1.xml.rels", PIVOT_TABLE_RELS_XML);
  zip.file("xl/pivotCache/pivotCacheDefinition1.xml", PIVOT_CACHE_DEFINITION_XML);
  zip.file("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", PIVOT_CACHE_DEFINITION_RELS_XML);
  zip.file("xl/pivotCache/pivotCacheRecords1.xml", PIVOT_CACHE_RECORDS_XML);

  // Workbook rels gain a `pivotCacheDefinition` relationship; we add
  // `rId99` as a fresh slot so we don't collide with whatever rIds the
  // base fixture already uses.
  const wbRelsPath = "xl/_rels/workbook.xml.rels";
  const wbRels = await zip.file(wbRelsPath)!.async("string");
  const augmentedWbRels = wbRels.replace(
    "</Relationships>",
    `<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>`
  );
  zip.file(wbRelsPath, augmentedWbRels);

  // workbook.xml grows a `<pivotCaches>` block right after `</sheets>`.
  const workbookXml = await zip.file("xl/workbook.xml")!.async("string");
  const augmentedWorkbookXml = workbookXml.replace(
    "</sheets>",
    `</sheets><pivotCaches><pivotCache cacheId="1" r:id="rId99"/></pivotCaches>`
  );
  zip.file("xl/workbook.xml", augmentedWorkbookXml);

  // The Inventory sheet gains a `pivotTable` rel and a `<pivotTableParts>`
  // child so Excel knows the table is anchored on this sheet.
  const sheetRelsPath = "xl/worksheets/_rels/sheet1.xml.rels";
  const existingSheetRels = zip.file(sheetRelsPath);
  const sheetRels = existingSheetRels
    ? await existingSheetRels.async("string")
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  const augmentedSheetRels = sheetRels.includes("</Relationships>")
    ? sheetRels.replace(
        "</Relationships>",
        `<Relationship Id="rIdPivot1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>`
      )
    : sheetRels;
  zip.file(sheetRelsPath, augmentedSheetRels);

  // [Content_Types].xml gets the three new overrides.
  const ctPath = "[Content_Types].xml";
  const ctXml = await zip.file(ctPath)!.async("string");
  const augmentedCt = ctXml.replace(
    "</Types>",
    `<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>` +
      `<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>` +
      `<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>` +
      `</Types>`
  );
  zip.file(ctPath, augmentedCt);

  const out = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return out;
}

describe("pivot tables — Phase 1 preservation", () => {
  it("lifts pivot parts into typed slots without losing bytes", async () => {
    const fixture = await buildPivotFixture();
    const snap = await parseXlsx(fixture);

    expect(snap.root.pivotTables).toHaveLength(1);
    expect(snap.root.pivotCaches).toHaveLength(1);

    const table = snap.root.pivotTables[0]!;
    expect(table.partPath).toBe("xl/pivotTables/pivotTable1.xml");
    expect(table.name).toBe("PivotInventory");
    expect(table.cacheId).toBe(1);
    expect(table.raw).toBe(PIVOT_TABLE_XML);
    expect(table.contentType).toBe(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
    );
    expect(table.relsXml).toBe(PIVOT_TABLE_RELS_XML);

    const cache = snap.root.pivotCaches[0]!;
    expect(cache.partPath).toBe("xl/pivotCache/pivotCacheDefinition1.xml");
    expect(cache.cacheId).toBe(1);
    expect(cache.raw).toBe(PIVOT_CACHE_DEFINITION_XML);
    expect(cache.recordsPartPath).toBe("xl/pivotCache/pivotCacheRecords1.xml");
    expect(cache.recordsRaw).toBe(PIVOT_CACHE_RECORDS_XML);
    expect(cache.relsXml).toBe(PIVOT_CACHE_DEFINITION_RELS_XML);
  });

  it("does not double-track pivot parts in opaqueParts", async () => {
    const fixture = await buildPivotFixture();
    const snap = await parseXlsx(fixture);

    const opaquePaths = [...snap.root.opaqueParts.keys()];
    expect(opaquePaths).not.toContain("xl/pivotTables/pivotTable1.xml");
    expect(opaquePaths).not.toContain("xl/pivotCache/pivotCacheDefinition1.xml");
    expect(opaquePaths).not.toContain("xl/pivotCache/pivotCacheRecords1.xml");
    expect(opaquePaths).not.toContain("xl/pivotTables/_rels/pivotTable1.xml.rels");
    expect(opaquePaths).not.toContain("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels");
  });

  it("round-trips pivot part bytes byte-identical via serialize → re-parse", async () => {
    const fixture = await buildPivotFixture();
    const snap = await parseXlsx(fixture);

    const written = await serializeXlsx(snap);
    const reparsed = await parseXlsx(new Uint8Array(written));

    // Pivot part hashes survive the round-trip identically.
    const pivotPartPaths = [
      "xl/pivotTables/pivotTable1.xml",
      "xl/pivotTables/_rels/pivotTable1.xml.rels",
      "xl/pivotCache/pivotCacheDefinition1.xml",
      "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
      "xl/pivotCache/pivotCacheRecords1.xml",
    ];
    for (const path of pivotPartPaths) {
      expect(reparsed.partHashes[path], path).toBe(snap.partHashes[path]);
    }

    // Typed slots are equivalent on both sides.
    expect(reparsed.root.pivotTables.map((p) => ({ partPath: p.partPath, name: p.name }))).toEqual(
      snap.root.pivotTables.map((p) => ({ partPath: p.partPath, name: p.name }))
    );
    expect(reparsed.root.pivotCaches.map((p) => p.cacheId)).toEqual(
      snap.root.pivotCaches.map((p) => p.cacheId)
    );

    // Hash via sha256 directly to guard against `partHashes` being
    // synthesized differently from the bytes we re-emit.
    const reparsedTable = reparsed.root.pivotTables[0]!;
    expect(sha256Hex(new TextEncoder().encode(reparsedTable.raw))).toBe(
      sha256Hex(new TextEncoder().encode(PIVOT_TABLE_XML))
    );
  });

  it("keeps source-range cell values unchanged across the round-trip", async () => {
    const fixture = await buildPivotFixture();
    const snap = await parseXlsx(fixture);

    const written = await serializeXlsx(snap);
    const reparsed = await parseXlsx(new Uint8Array(written));

    const inventory = reparsed.root.sheets.find((s) => s.name === "Inventory");
    expect(inventory).toBeDefined();
    const ws = reparsed.root.sheetjs.Sheets["Inventory"];
    const data = (ws as unknown as { "!data"?: Array<Array<{ v?: unknown }>> })["!data"]!;
    // Header row matches the original synthetic fixture
    // (Item / Qty / Unit price / Total).
    expect(data[0][0]?.v).toBe("Item");
    expect(data[0][1]?.v).toBe("Qty");
    // First data row remains 10 / 2.5 / 25 (per Widget).
    expect(data[1][0]?.v).toBe("Widget");
    expect(data[1][1]?.v).toBe(10);
  });
});
