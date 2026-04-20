/**
 * Mirror of `packages/docx/src/commands/embedded-roundtrip.test.ts`
 * for the PowerPoint side. See that file for the test recipe; the
 * only differences here are:
 *
 *   - We start from an existing fixture deck rather than a synthetic
 *     document, because PPTX has more required mandatory parts (slide
 *     master + layout + theme + slideId list) than DOCX and there's
 *     no `makeSyntheticPptx` helper.
 *   - Embedded chart workbooks live under `ppt/embeddings/`, not
 *     `word/embeddings/`.
 *   - Sheet refs use the same A1 grid layout as the DOCX side because
 *     `buildEmbeddedXlsx` is shared between the two formats.
 */

import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml } from "@officeai/core";
import JSZip from "jszip";
import { PptxAgent } from "../agent/agent.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, "04-multi-shape.pptx"));
  return PptxAgent.fromBuffer(buf);
}

async function reloadContainer(agent: PptxAgent): Promise<ooxml.OoxmlContainer> {
  const out = await agent.exportFile();
  return ooxml.OoxmlContainer.load(new Uint8Array(out));
}

function findEmbeddingPath(container: ooxml.OoxmlContainer): string {
  const paths = [...container.parts.keys()].filter((p) => p.startsWith("ppt/embeddings/"));
  expect(paths.length).toBeGreaterThan(0);
  return paths[0]!;
}

async function readSheetXml(container: ooxml.OoxmlContainer, partPath: string): Promise<string> {
  expect(container.has(partPath)).toBe(true);
  const inner = await JSZip.loadAsync(container.readBytes(partPath));
  const file = inner.file("xl/worksheets/sheet1.xml");
  expect(file).toBeTruthy();
  return file!.async("string");
}

function expectInlineString(sheetXml: string, ref: string, value: string): void {
  const re = new RegExp(`<c[^>]*r="${ref}"[^>]*t="inlineStr"[^>]*>\\s*<is>\\s*<t[^>]*>${value}</t>`);
  expect(sheetXml).toMatch(re);
}

function expectNumber(sheetXml: string, ref: string, value: number): void {
  const re = new RegExp(`<c[^>]*r="${ref}"(?![^>]*t="inlineStr")[^>]*>\\s*<v>${value}</v>`);
  expect(sheetXml).toMatch(re);
}

describe("pptx embedded roundtrip — chart workbook", () => {
  it("serialises insert-chart with the categories + series projected into the embedded sheet", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "pptx:insert-chart",
      payload: {
        slideIndex: 0,
        x: 1_000_000,
        y: 1_000_000,
        chartType: "bar",
        title: "Sales",
        categories: ["Q1", "Q2", "Q3"],
        series: [
          { name: "EU", values: [10, 20, 30] },
          { name: "US", values: [15, 25, 35] },
        ],
      },
    });

    const reloaded = await reloadContainer(agent);
    const sheetXml = await readSheetXml(reloaded, findEmbeddingPath(reloaded));

    expectInlineString(sheetXml, "B1", "EU");
    expectInlineString(sheetXml, "C1", "US");
    expectInlineString(sheetXml, "A2", "Q1");
    expectInlineString(sheetXml, "A3", "Q2");
    expectInlineString(sheetXml, "A4", "Q3");
    expectNumber(sheetXml, "B2", 10);
    expectNumber(sheetXml, "B3", 20);
    expectNumber(sheetXml, "B4", 30);
    expectNumber(sheetXml, "C2", 15);
    expectNumber(sheetXml, "C3", 25);
    expectNumber(sheetXml, "C4", 35);
  });

  it("set-chart-data refreshes the embedded sheet after insert", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "pptx:insert-chart",
      payload: {
        slideIndex: 0,
        x: 0,
        y: 0,
        chartType: "line",
        categories: ["Old"],
        series: [{ name: "S0", values: [1] }],
      },
    });
    const snap0 = agent.getSnapshot();
    const slide0 = snap0.root.slides[0]!;
    const chartShape = slide0.shapes.find((s) => s.kind === "chart");
    expect(chartShape).toBeDefined();

    await agent.applyCommand({
      type: "pptx:set-chart-data",
      payload: {
        slideIndex: 0,
        shapeId: chartShape!.id,
        categories: ["Mon", "Tue"],
        series: [{ name: "Updated", values: [42, 43] }],
      },
    });

    const reloaded = await reloadContainer(agent);
    const sheetXml = await readSheetXml(reloaded, findEmbeddingPath(reloaded));

    expectInlineString(sheetXml, "B1", "Updated");
    expectInlineString(sheetXml, "A2", "Mon");
    expectInlineString(sheetXml, "A3", "Tue");
    expectNumber(sheetXml, "B2", 42);
    expectNumber(sheetXml, "B3", 43);
    expect(sheetXml).not.toMatch(/<t[^>]*>S0<\/t>/);
  });
});

describe("pptx embedded roundtrip — OLE spreadsheet", () => {
  it("serialises insert-spreadsheet with the supplied grid in the embedded sheet", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "pptx:insert-spreadsheet",
      payload: {
        slideIndex: 0,
        x: 500_000,
        y: 500_000,
        data: [
          ["Name", "Score"],
          ["Ada", 91],
          ["Linus", 88],
        ],
      },
    });

    const reloaded = await reloadContainer(agent);
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
      type: "pptx:insert-spreadsheet",
      payload: {
        slideIndex: 0,
        x: 0,
        y: 0,
        data: [["seed"]],
      },
    });
    // Roundtrip once so the agent's snapshot embeddings carry real
    // bytes (insert seeds `pendingGrid`); we want to exercise the
    // "replace bytes" path explicitly, mirroring the editor's
    // double-click → edit → save flow.
    const firstOut = await agent.exportFile();
    const reagent = await PptxAgent.fromBuffer(new Uint8Array(firstOut));
    const firstReload = await ooxml.OoxmlContainer.load(new Uint8Array(firstOut));
    const embeddingPath = findEmbeddingPath(firstReload);

    // Build replacement bytes via a fresh insert-spreadsheet on a
    // disposable agent + serialise so we stay decoupled from the
    // exact wire format of `buildEmbeddedXlsx`.
    const sourceAgent = await loadAgent();
    await sourceAgent.applyCommand({
      type: "pptx:insert-spreadsheet",
      payload: {
        slideIndex: 0,
        x: 0,
        y: 0,
        data: [
          ["Replaced", "Header"],
          ["row1", 7],
        ],
      },
    });
    const sourceOut = await sourceAgent.exportFile();
    const sourceReload = await ooxml.OoxmlContainer.load(new Uint8Array(sourceOut));
    const replacementBytes = sourceReload.readBytes(findEmbeddingPath(sourceReload));

    await reagent.applyCommand({
      type: "pptx:update-spreadsheet",
      payload: {
        embeddingPartPath: embeddingPath,
        bytes: replacementBytes,
      },
    });

    const finalReload = await reloadContainer(reagent);
    const sheetXml = await readSheetXml(finalReload, embeddingPath);
    expectInlineString(sheetXml, "A1", "Replaced");
    expectInlineString(sheetXml, "B1", "Header");
    expectInlineString(sheetXml, "A2", "row1");
    expectNumber(sheetXml, "B2", 7);
    expect(sheetXml).not.toMatch(/<t[^>]*>seed<\/t>/);
  });
});
