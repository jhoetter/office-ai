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

describe("xlsx:set-page-setup", () => {
  it("writes orientation + paperSize attributes onto <pageSetup>", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;

    await agent.applyCommand({
      type: "xlsx:set-page-setup",
      payload: { sheet: sheetName, orientation: "landscape", paperSize: 9 },
    });

    const sheet = agent.getSnapshot().root.sheets.find((s) => s.name === sheetName)!;
    expect(sheet.pageSetupXml).toBeDefined();
    expect(sheet.pageSetupXml).toMatch(/orientation="landscape"/);
    expect(sheet.pageSetupXml).toMatch(/paperSize="9"/);
  });

  it("rejects scale outside 10–400", async () => {
    const agent = await loadAgent();
    const result = await agent.applyCommand({
      type: "xlsx:set-page-setup",
      payload: { sheet: agent.listSheets()[0]!.name, scale: 5 },
    });
    expect(result.status).toBe("rejected");
    if (result.status !== "rejected") return;
    expect(result.rejection.message).toMatch(/scale/);
  });

  it("clears the <pageSetup> element when clear:true", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-page-setup",
      payload: { sheet: sheetName, orientation: "portrait" },
    });
    await agent.applyCommand({
      type: "xlsx:set-page-setup",
      payload: { sheet: sheetName, clear: true },
    });
    const sheet = agent.getSnapshot().root.sheets.find((s) => s.name === sheetName)!;
    expect(sheet.pageSetupXml).toBeUndefined();
  });

  it("survives a save/reload round-trip", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-page-setup",
      payload: { sheet: sheetName, orientation: "landscape", fitToWidth: 1, fitToHeight: 0 },
    });
    const bytes = await agent.exportFile();
    const reopened = await XlsxAgent.fromBuffer(bytes);
    const sheet = reopened.getSnapshot().root.sheets.find((s) => s.name === sheetName)!;
    expect(sheet.pageSetupXml).toMatch(/orientation="landscape"/);
    expect(sheet.pageSetupXml).toMatch(/fitToWidth="1"/);
  });
});

describe("xlsx:set-page-margins", () => {
  it("applies the wide preset", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-page-margins",
      payload: { sheet: sheetName, preset: "wide" },
    });
    const sheet = agent.getSnapshot().root.sheets.find((s) => s.name === sheetName)!;
    expect(sheet.pageMarginsXml).toMatch(/left="1"/);
    expect(sheet.pageMarginsXml).toMatch(/right="1"/);
    expect(sheet.pageMarginsXml).toMatch(/top="1"/);
  });

  it("merges individual overrides on top of preset", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-page-margins",
      payload: { sheet: sheetName, preset: "narrow", topIn: 0.9 },
    });
    const sheet = agent.getSnapshot().root.sheets.find((s) => s.name === sheetName)!;
    expect(sheet.pageMarginsXml).toMatch(/left="0\.25"/);
    expect(sheet.pageMarginsXml).toMatch(/top="0\.9"/);
  });

  it("rejects negative margins", async () => {
    const agent = await loadAgent();
    const result = await agent.applyCommand({
      type: "xlsx:set-page-margins",
      payload: { sheet: agent.listSheets()[0]!.name, leftIn: -1 },
    });
    expect(result.status).toBe("rejected");
  });
});

describe("xlsx:set-print-options", () => {
  it("toggles individual flags", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-print-options",
      payload: { sheet: sheetName, gridLines: true, headings: true },
    });
    const sheet = agent.getSnapshot().root.sheets.find((s) => s.name === sheetName)!;
    expect(sheet.printOptionsXml).toMatch(/gridLines="1"/);
    expect(sheet.printOptionsXml).toMatch(/headings="1"/);
  });

  it("removes attribute when toggled false", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-print-options",
      payload: { sheet: sheetName, gridLines: true },
    });
    await agent.applyCommand({
      type: "xlsx:set-print-options",
      payload: { sheet: sheetName, gridLines: false },
    });
    const sheet = agent.getSnapshot().root.sheets.find((s) => s.name === sheetName)!;
    expect(sheet.printOptionsXml ?? "").not.toMatch(/gridLines/);
  });
});

describe("xlsx:set-print-area / xlsx:set-print-titles", () => {
  it("writes a sheet-scoped _xlnm.Print_Area defined name", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-print-area",
      payload: { sheet: sheetName, range: "A1:D20" },
    });
    const dn = agent
      .getSnapshot()
      .root.definedNames.find((d) => d.name === "_xlnm.Print_Area" && d.scope === sheetName);
    expect(dn).toBeDefined();
    expect(dn?.refersTo).toBe(`'${sheetName}'!$A$1:$D$20`);
  });

  it("clears the print area on demand", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-print-area",
      payload: { sheet: sheetName, range: "A1:B2" },
    });
    await agent.applyCommand({
      type: "xlsx:set-print-area",
      payload: { sheet: sheetName, clear: true },
    });
    const dn = agent
      .getSnapshot()
      .root.definedNames.find((d) => d.name === "_xlnm.Print_Area" && d.scope === sheetName);
    expect(dn).toBeUndefined();
  });

  it("writes Print_Titles with rows-only", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-print-titles",
      payload: { sheet: sheetName, rows: "1:2" },
    });
    const dn = agent
      .getSnapshot()
      .root.definedNames.find((d) => d.name === "_xlnm.Print_Titles" && d.scope === sheetName);
    expect(dn?.refersTo).toBe(`'${sheetName}'!$1:$2`);
  });

  it("writes Print_Titles with rows + cols", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-print-titles",
      payload: { sheet: sheetName, rows: "1", cols: "A:B" },
    });
    const dn = agent
      .getSnapshot()
      .root.definedNames.find((d) => d.name === "_xlnm.Print_Titles" && d.scope === sheetName);
    expect(dn?.refersTo).toBe(`'${sheetName}'!$1:$1,'${sheetName}'!$A:$B`);
  });

  it("survives a save/reload round-trip", async () => {
    const agent = await loadAgent();
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-print-area",
      payload: { sheet: sheetName, range: "A1:Z100" },
    });
    await agent.applyCommand({
      type: "xlsx:set-print-titles",
      payload: { sheet: sheetName, rows: "1:1" },
    });
    const bytes = await agent.exportFile();
    const reopened = await XlsxAgent.fromBuffer(bytes);
    const dnArea = reopened
      .getSnapshot()
      .root.definedNames.find((d) => d.name === "_xlnm.Print_Area" && d.scope === sheetName);
    const dnTitles = reopened
      .getSnapshot()
      .root.definedNames.find((d) => d.name === "_xlnm.Print_Titles" && d.scope === sheetName);
    expect(dnArea?.refersTo).toBe(`'${sheetName}'!$A$1:$Z$100`);
    expect(dnTitles?.refersTo).toBe(`'${sheetName}'!$1:$1`);
  });
});
