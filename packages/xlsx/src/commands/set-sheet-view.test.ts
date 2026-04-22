import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { XlsxAgent } from "../agent/agent.js";

const here = dirname(fileURLToPath(import.meta.url));
const fixtures = resolve(here, "../../../../fixtures/xlsx/synthetic");

async function loadAgent(name = "02-multi-sheet.xlsx"): Promise<XlsxAgent> {
  return XlsxAgent.fromBuffer(new Uint8Array(await readFile(resolve(fixtures, name))));
}

function viewXml(agent: XlsxAgent, sheetName: string): string {
  return agent.getSnapshot().root.sheets.find((s) => s.name === sheetName)!.sheetViewsXml ?? "";
}

describe("xlsx:set-sheet-view", () => {
  it("toggles showGridLines off (writes \"0\")", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;

    await agent.applyCommand({
      type: "xlsx:set-sheet-view",
      payload: { sheet: sheetName, showGridLines: false },
    });

    expect(viewXml(agent, sheetName)).toMatch(/showGridLines="0"/);
  });

  it("dropping showGridLines back to true removes the attribute", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;

    await agent.applyCommand({
      type: "xlsx:set-sheet-view",
      payload: { sheet: sheetName, showGridLines: false },
    });
    await agent.applyCommand({
      type: "xlsx:set-sheet-view",
      payload: { sheet: sheetName, showGridLines: true },
    });

    expect(viewXml(agent, sheetName)).not.toMatch(/showGridLines/);
  });

  it("sets view='pageLayout' and zoomScale", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;

    await agent.applyCommand({
      type: "xlsx:set-sheet-view",
      payload: { sheet: sheetName, view: "pageLayout", zoomScale: 125 },
    });

    const xml = viewXml(agent, sheetName);
    expect(xml).toMatch(/view="pageLayout"/);
    expect(xml).toMatch(/zoomScale="125"/);
  });

  it("setting view='normal' drops the attribute (Excel default)", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-sheet-view",
      payload: { sheet: sheetName, view: "pageLayout" },
    });
    await agent.applyCommand({
      type: "xlsx:set-sheet-view",
      payload: { sheet: sheetName, view: "normal" },
    });
    expect(viewXml(agent, sheetName)).not.toMatch(/view=/);
  });

  it("rejects an empty payload", async () => {
    const agent = await loadAgent();
    const result = await agent.applyCommand({
      type: "xlsx:set-sheet-view",
      payload: { sheet: agent.listSheets()[0]!.name },
    });
    expect(result.status).toBe("rejected");
    if (result.status !== "rejected") return;
    expect(result.rejection.message).toMatch(/at least one/);
  });

  it("rejects out-of-range zoomScale", async () => {
    const agent = await loadAgent();
    const result = await agent.applyCommand({
      type: "xlsx:set-sheet-view",
      payload: { sheet: agent.listSheets()[0]!.name, zoomScale: 1000 },
    });
    expect(result.status).toBe("rejected");
  });

  it("survives a save/reload round-trip", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-sheet-view",
      payload: { sheet: sheetName, view: "pageBreakPreview", showRowColHeaders: false, zoomScale: 75 },
    });
    const bytes = await agent.exportFile();
    const reopened = await XlsxAgent.fromBuffer(bytes);
    const xml = reopened.getSnapshot().root.sheets.find((s) => s.name === sheetName)!.sheetViewsXml ?? "";
    expect(xml).toMatch(/view="pageBreakPreview"/);
    expect(xml).toMatch(/showRowColHeaders="0"/);
    expect(xml).toMatch(/zoomScale="75"/);
  });
});
