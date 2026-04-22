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

describe("xlsx:set-calc-mode", () => {
  it("sets calcMode=manual on <calcPr>", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "xlsx:set-calc-mode",
      payload: { calcMode: "manual" },
    });
    const xml = agent.getSnapshot().root.calcPrXml;
    expect(xml).toMatch(/calcMode="manual"/);
  });

  it("removes calcMode attr when switching back to auto", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "xlsx:set-calc-mode",
      payload: { calcMode: "manual" },
    });
    await agent.applyCommand({
      type: "xlsx:set-calc-mode",
      payload: { calcMode: "auto" },
    });
    const xml = agent.getSnapshot().root.calcPrXml ?? "";
    expect(xml).not.toMatch(/calcMode=/);
  });

  it("toggles iterative calc + iterateCount", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "xlsx:set-calc-mode",
      payload: { iterate: true, iterateCount: 50, iterateDelta: 0.001 },
    });
    const xml = agent.getSnapshot().root.calcPrXml ?? "";
    expect(xml).toMatch(/iterate="1"/);
    expect(xml).toMatch(/iterateCount="50"/);
    expect(xml).toMatch(/iterateDelta="0.001"/);
  });

  it("survives a save/reload round-trip", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "xlsx:set-calc-mode",
      payload: { calcMode: "manual", calcOnSave: true },
    });
    const bytes = await agent.exportFile();
    const reopened = await XlsxAgent.fromBuffer(bytes);
    const xml = reopened.getSnapshot().root.calcPrXml ?? "";
    expect(xml).toMatch(/calcMode="manual"/);
    expect(xml).toMatch(/calcOnSave="1"/);
  });
});
