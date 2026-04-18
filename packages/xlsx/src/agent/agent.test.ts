import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { deterministicIdMinter, ooxml, sha256Hex } from "@officeai/core";
import { XlsxAgent } from "./agent.js";
import { parseXlsx } from "../parser/parse.js";

const here = dirname(fileURLToPath(import.meta.url));
const fixtures = resolve(here, "../../../../fixtures/xlsx/synthetic");

async function loadFixture(name: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(resolve(fixtures, name)));
}

describe("XlsxAgent end-to-end", () => {
  it("trivial-edit roundtrip preserves untouched parts byte-for-byte", async () => {
    const buf = await loadFixture("02-multi-sheet.xlsx");
    const original = await ooxml.OoxmlContainer.load(buf);
    const originalHashes: Record<string, string> = {};
    for (const path of original.parts.keys()) {
      originalHashes[path] = sha256Hex(original.readBytes(path));
    }

    const agent = await XlsxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const sheets = agent.listSheets();
    const target = sheets[0]!.name;
    const targetPart = `xl/worksheets/sheet1.xml`;

    await agent.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet: target, ref: "Z99", value: 42 },
    });
    const out = await agent.exportFile();
    const reloaded = await ooxml.OoxmlContainer.load(out);

    for (const path of reloaded.parts.keys()) {
      if (path === targetPart) continue;
      const before = originalHashes[path];
      if (before === undefined) continue;
      const after = sha256Hex(reloaded.readBytes(path));
      expect(after, `untouched part ${path} should be byte-identical`).toBe(before);
    }
  });

  it("re-parses an exported edited file with the new value in place", async () => {
    const buf = await loadFixture("01-single-sheet-numbers.xlsx");
    const agent = await XlsxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet: sheetName, ref: "AA50", value: "agent-edit" },
    });
    const out = await agent.exportFile();
    const reparsed = await parseXlsx(out, { idMinter: deterministicIdMinter("x") });
    const sheet = reparsed.root.sheets[0];
    const matched = [...sheet.cells.values()].find((c) => c.value === "agent-edit");
    expect(matched).toBeDefined();
    expect(matched!.row).toBe(49);
    expect(matched!.col).toBe(26);
  });

  it("listSheets returns names + indexes in tab order", async () => {
    const buf = await loadFixture("02-multi-sheet.xlsx");
    const agent = await XlsxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const list = agent.listSheets();
    expect(list.length).toBeGreaterThan(1);
    expect(list.every((s) => s.kind === "worksheet" || s.kind === "chartsheet")).toBe(true);
    const indexes = list.map((s) => s.index);
    expect(indexes).toEqual([...indexes].sort((a, b) => a - b));
  });

  it("toMarkdown includes a header for each worksheet and renders cells", async () => {
    const buf = await loadFixture("01-single-sheet-numbers.xlsx");
    const agent = await XlsxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const sheetName = agent.listSheets()[0]!.name;
    const md = agent.toMarkdown({ maxRows: 5, maxCols: 5 });
    expect(md).toContain(`## ${sheetName}`);
    expect(md).toMatch(/\| ---/);
  });

  it("getRange returns the sparse cells inside a rectangle", async () => {
    const buf = await loadFixture("01-single-sheet-numbers.xlsx");
    const agent = await XlsxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-range-values",
      payload: {
        sheet: sheetName,
        range: "AA1:AB2",
        values: [
          [1, 2],
          [3, 4],
        ],
      },
    });
    const snap = agent.getRange({ kind: "xlsx-range", sheet: sheetName, range: "AA1:AB2" });
    expect(snap.rows).toBe(2);
    expect(snap.cols).toBe(2);
    expect(snap.cells.map((c) => c.value)).toEqual([1, 2, 3, 4]);
    expect(snap.cells.map((c) => c.ref)).toEqual(["AA1", "AB1", "AA2", "AB2"]);
  });

  it("search finds substrings inside string cells across sheets", async () => {
    const buf = await loadFixture("02-multi-sheet.xlsx");
    const agent = await XlsxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet: sheetName, ref: "Z1", value: "needle in haystack" },
    });
    const results = agent.search({ query: "needle" });
    expect(results.length).toBe(1);
    expect(results[0].sheet).toBe(sheetName);
    expect(results[0].ref).toBe("Z1");
    expect(results[0].match).toBe("needle");
  });

  it("getDiff between two snapshots reports cell + sheet changes", async () => {
    const buf = await loadFixture("02-multi-sheet.xlsx");
    const agent = await XlsxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const before = agent.getSnapshot();
    const sheetName = agent.listSheets()[0]!.name;
    await agent.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet: sheetName, ref: "AA1", value: "new" },
    });
    await agent.applyCommand({
      type: "xlsx:rename-sheet",
      payload: { name: sheetName, newName: "Renamed" },
    });
    const after = agent.getSnapshot();
    const diff = agent.getDiff(before, after);
    expect(diff.format).toBe("xlsx");
    expect(diff.changes.length).toBeGreaterThanOrEqual(2);
    const summaries = diff.changes.map((c) => c.summary);
    expect(summaries.some((s) => s?.includes("AA1"))).toBe(true);
    expect(summaries.some((s) => s?.includes("Renamed"))).toBe(true);
  });

  it("agent-sourced commands stage as pending and require approval", async () => {
    const buf = await loadFixture("01-single-sheet-numbers.xlsx");
    const agent = await XlsxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const sheetName = agent.listSheets()[0]!.name;
    const m = await agent.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet: sheetName, ref: "AA1", value: 7 },
      source: "agent",
    });
    expect(m.status).toBe("pending");
    expect(agent.getPendingMutations()).toHaveLength(1);
    agent.approveMutation(m.id);
    expect(agent.getPendingMutations()).toHaveLength(0);
  });

  it("is fully headless (works without DOM globals)", async () => {
    expect(typeof globalThis.window).toBe("undefined");
    expect(typeof globalThis.document).toBe("undefined");
    const buf = await loadFixture("01-single-sheet-numbers.xlsx");
    const agent = await XlsxAgent.fromBuffer(buf);
    expect(agent.getSnapshot().format).toBe("xlsx");
  });
});
