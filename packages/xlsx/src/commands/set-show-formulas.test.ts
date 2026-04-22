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

describe("xlsx:set-show-formulas", () => {
  it("stamps showFormulas=1 onto the primary sheetView", async () => {
    const agent = await loadAgent();
    const target = agent.listSheets()[0]!.name;

    await agent.applyCommand({
      type: "xlsx:set-show-formulas",
      payload: { sheet: target, show: true },
    });

    const sheet = agent.getSnapshot().root.sheets.find((s) => s.name === target)!;
    expect(sheet.sheetViewsXml ?? "").toMatch(/showFormulas="1"/);
  });

  it("removes the attribute when toggled back off", async () => {
    const agent = await loadAgent();
    const target = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-show-formulas",
      payload: { sheet: target, show: true },
    });
    await agent.applyCommand({
      type: "xlsx:set-show-formulas",
      payload: { sheet: target, show: false },
    });
    const sheet = agent.getSnapshot().root.sheets.find((s) => s.name === target)!;
    expect(sheet.sheetViewsXml ?? "").not.toMatch(/showFormulas=/);
  });

  it("survives a save/reload round-trip", async () => {
    const agent = await loadAgent();
    const target = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-show-formulas",
      payload: { sheet: target, show: true },
    });
    const bytes = await agent.exportFile();
    const reopened = await XlsxAgent.fromBuffer(bytes);
    const sheet = reopened.getSnapshot().root.sheets.find((s) => s.name === target)!;
    expect(sheet.sheetViewsXml ?? "").toMatch(/showFormulas="1"/);
  });
});
