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

describe("xlsx:set-sheet-protection", () => {
  it("stamps a <sheetProtection> element on the sheet model", async () => {
    const agent = await loadAgent();
    const target = agent.listSheets()[0]!.name;

    await agent.applyCommand({
      type: "xlsx:set-sheet-protection",
      payload: { sheet: target, enabled: true, formatCells: false, autoFilter: true },
    });

    const sheet = agent.getSnapshot().root.sheets.find((s) => s.name === target)!;
    expect(sheet.sheetProtectionXml).toMatch(/<sheetProtection\b/);
    expect(sheet.sheetProtectionXml).toMatch(/sheet="1"/);
    expect(sheet.sheetProtectionXml).toMatch(/formatCells="0"/);
    expect(sheet.sheetProtectionXml).toMatch(/autoFilter="1"/);
  });

  it("clears the element on disable", async () => {
    const agent = await loadAgent();
    const target = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-sheet-protection",
      payload: { sheet: target, enabled: true },
    });
    await agent.applyCommand({
      type: "xlsx:set-sheet-protection",
      payload: { sheet: target, enabled: false },
    });
    const sheet = agent.getSnapshot().root.sheets.find((s) => s.name === target)!;
    expect(sheet.sheetProtectionXml).toBeUndefined();
  });

  it("rejects plaintext passwords on the payload", async () => {
    const agent = await loadAgent();
    const target = agent.listSheets()[0]!.name;
    const result = await agent.applyCommand({
      type: "xlsx:set-sheet-protection",
      payload: { sheet: target, enabled: true, password: "secret" } as never,
    });
    expect(result.status).toBe("rejected");
    if (result.status !== "rejected") return;
    expect(result.rejection.message).toMatch(/Plaintext passwords/);
  });

  it("survives a save/reload round-trip", async () => {
    const agent = await loadAgent();
    const target = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-sheet-protection",
      payload: { sheet: target, enabled: true, formatCells: false },
    });
    const bytes = await agent.exportFile();
    const reopened = await XlsxAgent.fromBuffer(bytes);
    const sheet = reopened.getSnapshot().root.sheets.find((s) => s.name === target)!;
    expect(sheet.sheetProtectionXml).toMatch(/<sheetProtection\b/);
    expect(sheet.sheetProtectionXml).toMatch(/formatCells="0"/);
  });
});

describe("xlsx:set-workbook-protection", () => {
  it("stamps a <workbookProtection> element on the workbook model", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "xlsx:set-workbook-protection",
      payload: { enabled: true, lockStructure: true, lockWindows: false },
    });
    const wb = agent.getSnapshot().root;
    expect(wb.workbookProtectionXml).toMatch(/<workbookProtection\b/);
    expect(wb.workbookProtectionXml).toMatch(/lockStructure="1"/);
  });

  it("clears the element on disable", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "xlsx:set-workbook-protection",
      payload: { enabled: true, lockStructure: true },
    });
    await agent.applyCommand({
      type: "xlsx:set-workbook-protection",
      payload: { enabled: false },
    });
    expect(agent.getSnapshot().root.workbookProtectionXml).toBeUndefined();
  });

  it("survives a save/reload round-trip", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "xlsx:set-workbook-protection",
      payload: { enabled: true, lockStructure: true, lockWindows: true },
    });
    const bytes = await agent.exportFile();
    const reopened = await XlsxAgent.fromBuffer(bytes);
    const wb = reopened.getSnapshot().root;
    expect(wb.workbookProtectionXml).toMatch(/lockStructure="1"/);
    expect(wb.workbookProtectionXml).toMatch(/lockWindows="1"/);
  });
});
