/**
 * End-to-end roundtrip tests for the embedded-spreadsheet / chart
 * authoring + edit pipeline. Each test:
 *
 *   1. Parses a synthetic DOCX into an agent.
 *   2. Dispatches one or more bus commands that mutate the embedded
 *      `.xlsx` payload (insert-chart, insert-spreadsheet,
 *      update-spreadsheet, set-chart-data).
 *   3. Serialises the agent back into a `.docx` package.
 *   4. Reloads the package via {@link ooxml.OoxmlContainer}.
 *   5. Unzips the embedded `.xlsx` part and asserts cell-level
 *      values are present in `xl/worksheets/sheet1.xml`.
 *
 * Cell-level assertions read the raw worksheet XML rather than going
 * through `parseXlsx` because the embedded payload is intentionally
 * minimalist — `buildEmbeddedXlsx` only emits inline strings + numeric
 * `<c><v>` cells with no shared-strings table — and the regex checks
 * make the failure surface much sharper when something downstream
 * accidentally drops a cell or rewrites the column letters.
 */

import { describe, expect, it } from "vitest";
import { deterministicIdMinter, ooxml } from "@officeai/core";
import JSZip from "jszip";
import { DocxAgent } from "../agent/agent.js";
import { serializeDocx } from "../serializer/serialize.js";
import { DEFAULT_DOC_ROOT_ATTRS, makeSyntheticDocx } from "../test-utils/synthetic.js";

function plainDoc(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:p><w:r><w:t>placeholder</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
}

async function loadAgent(): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: plainDoc() });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

async function readEmbedded(container: ooxml.OoxmlContainer, partPath: string): Promise<JSZip> {
  expect(container.has(partPath)).toBe(true);
  const bytes = container.readBytes(partPath);
  expect(bytes.byteLength).toBeGreaterThan(0);
  return JSZip.loadAsync(bytes);
}

async function readSheetXml(container: ooxml.OoxmlContainer, partPath: string): Promise<string> {
  const inner = await readEmbedded(container, partPath);
  const file = inner.file("xl/worksheets/sheet1.xml");
  expect(file).toBeTruthy();
  return file!.async("string");
}

function findEmbeddingPath(container: ooxml.OoxmlContainer): string {
  const paths = [...container.parts.keys()].filter((p) => p.startsWith("word/embeddings/"));
  expect(paths.length).toBeGreaterThan(0);
  return paths[0]!;
}

/** Match an inline-string `<c r="A1" t="inlineStr"><is><t...>VAL</t></is></c>`. */
function expectInlineString(sheetXml: string, ref: string, value: string): void {
  const re = new RegExp(`<c[^>]*r="${ref}"[^>]*t="inlineStr"[^>]*>\\s*<is>\\s*<t[^>]*>${value}</t>`);
  expect(sheetXml).toMatch(re);
}

/** Match a numeric `<c r="A1"><v>VAL</v></c>` cell (no `t` attr or `t="n"`). */
function expectNumber(sheetXml: string, ref: string, value: number): void {
  const re = new RegExp(`<c[^>]*r="${ref}"(?![^>]*t="inlineStr")[^>]*>\\s*<v>${value}</v>`);
  expect(sheetXml).toMatch(re);
}

describe("docx embedded roundtrip — chart workbook", () => {
  it("serialises insert-chart with the categories + series projected into the embedded sheet", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "docx:insert-chart",
      payload: {
        at: { paragraph: 0 },
        chartType: "bar",
        title: "Revenue",
        categories: ["Jan", "Feb", "Mar"],
        series: [
          { name: "EU", values: [10, 20, 30] },
          { name: "US", values: [40, 50, 60] },
        ],
      },
      source: "human",
    });

    const out = await serializeDocx(agent.getSnapshot());
    const reloaded = await ooxml.OoxmlContainer.load(out);

    const sheetXml = await readSheetXml(reloaded, findEmbeddingPath(reloaded));

    // Header row (B1, C1) holds the series names; A1 stays empty.
    expectInlineString(sheetXml, "B1", "EU");
    expectInlineString(sheetXml, "C1", "US");
    // Category column.
    expectInlineString(sheetXml, "A2", "Jan");
    expectInlineString(sheetXml, "A3", "Feb");
    expectInlineString(sheetXml, "A4", "Mar");
    // Series values.
    expectNumber(sheetXml, "B2", 10);
    expectNumber(sheetXml, "B3", 20);
    expectNumber(sheetXml, "B4", 30);
    expectNumber(sheetXml, "C2", 40);
    expectNumber(sheetXml, "C3", 50);
    expectNumber(sheetXml, "C4", 60);
  });

  it("set-chart-data refreshes the embedded sheet after insert", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "docx:insert-chart",
      payload: {
        at: { paragraph: 0 },
        chartType: "line",
        categories: ["A"],
        series: [{ name: "Old", values: [1] }],
      },
      source: "human",
    });
    await agent.applyCommand({
      type: "docx:set-chart-data",
      payload: {
        chartPartPath: "word/charts/chart1.xml",
        categories: ["X", "Y"],
        series: [{ name: "New", values: [99, 100] }],
      },
      source: "human",
    });

    const out = await serializeDocx(agent.getSnapshot());
    const reloaded = await ooxml.OoxmlContainer.load(out);
    const sheetXml = await readSheetXml(reloaded, findEmbeddingPath(reloaded));

    expectInlineString(sheetXml, "B1", "New");
    expectInlineString(sheetXml, "A2", "X");
    expectInlineString(sheetXml, "A3", "Y");
    expectNumber(sheetXml, "B2", 99);
    expectNumber(sheetXml, "B3", 100);
    // Stale data must be gone.
    expect(sheetXml).not.toMatch(/<t[^>]*>Old<\/t>/);
  });
});

describe("docx embedded roundtrip — OLE spreadsheet", () => {
  it("serialises insert-spreadsheet with the supplied grid in the embedded sheet", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "docx:insert-spreadsheet",
      payload: {
        at: { paragraph: 0 },
        data: [
          ["Name", "Score"],
          ["Ada", 91],
          ["Linus", 88],
        ],
      },
      source: "human",
    });

    const out = await serializeDocx(agent.getSnapshot());
    const reloaded = await ooxml.OoxmlContainer.load(out);
    const sheetXml = await readSheetXml(reloaded, findEmbeddingPath(reloaded));

    expectInlineString(sheetXml, "A1", "Name");
    expectInlineString(sheetXml, "B1", "Score");
    expectInlineString(sheetXml, "A2", "Ada");
    expectInlineString(sheetXml, "A3", "Linus");
    expectNumber(sheetXml, "B2", 91);
    expectNumber(sheetXml, "B3", 88);
  });

  it("update-spreadsheet swaps the embedded bytes wholesale", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "docx:insert-spreadsheet",
      payload: {
        at: { paragraph: 0 },
        data: [["seed"]],
      },
      source: "human",
    });

    // Serialise once to materialise the embedded xlsx, then mutate
    // those bytes via update-spreadsheet so we exercise the
    // "agent has real bytes, not pendingGrid" path the editor uses
    // after a double-click → edit → save round-trip.
    const firstOut = await serializeDocx(agent.getSnapshot());
    const firstReload = await ooxml.OoxmlContainer.load(firstOut);
    const embeddingPath = findEmbeddingPath(firstReload);

    // Re-parse so the agent's snapshot.embeddings carries real bytes
    // (insert-spreadsheet seeds `pendingGrid`; we want to test the
    // "replace bytes" path explicitly).
    const reagent = await DocxAgent.fromBuffer(firstOut, {
      idMinter: deterministicIdMinter(),
    });

    // Build replacement bytes via a fresh insert-spreadsheet on a
    // disposable agent + serialise, so the test stays decoupled from
    // `buildEmbeddedXlsx`'s exact wire format.
    const sourceAgent = await loadAgent();
    await sourceAgent.applyCommand({
      type: "docx:insert-spreadsheet",
      payload: {
        at: { paragraph: 0 },
        data: [
          ["Replaced", "Header"],
          ["row1", 7],
        ],
      },
      source: "human",
    });
    const sourcePackage = await serializeDocx(sourceAgent.getSnapshot());
    const sourceReload = await ooxml.OoxmlContainer.load(sourcePackage);
    const replacementBytes = sourceReload.readBytes(findEmbeddingPath(sourceReload));

    await reagent.applyCommand({
      type: "docx:update-spreadsheet",
      payload: {
        embeddingPartPath: embeddingPath,
        bytes: replacementBytes,
      },
      source: "human",
    });

    const out = await serializeDocx(reagent.getSnapshot());
    const reloaded = await ooxml.OoxmlContainer.load(out);
    const inner = await readEmbedded(reloaded, embeddingPath);
    const sheet = await inner.file("xl/worksheets/sheet1.xml")!.async("string");

    expectInlineString(sheet, "A1", "Replaced");
    expectInlineString(sheet, "B1", "Header");
    expectInlineString(sheet, "A2", "row1");
    expectNumber(sheet, "B2", 7);
    expect(sheet).not.toMatch(/<t[^>]*>seed<\/t>/);
  });
});
