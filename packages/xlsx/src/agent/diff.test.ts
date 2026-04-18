import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { XlsxAgent } from "./agent.js";
import { diffXlsxSnapshots } from "./diff.js";

const here = dirname(fileURLToPath(import.meta.url));
const fixtures = resolve(here, "../../../../fixtures/xlsx/synthetic");

async function loadAgent(name: string): Promise<XlsxAgent> {
  const buf = new Uint8Array(await readFile(resolve(fixtures, name)));
  return XlsxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

describe("diffXlsxSnapshots (structural snapshot diff)", () => {
  it("reports an inserted cell as node-inserted", async () => {
    const a = await loadAgent("01-single-sheet-numbers.xlsx");
    const sheetName = a.listSheets()[0]!.name;
    const before = a.getSnapshot();
    await a.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet: sheetName, ref: "Z99", value: "fresh" },
    });
    const after = a.getSnapshot();
    const diff = diffXlsxSnapshots(before, after);
    const inserts = diff.changes.filter((c) => c.kind === "node-inserted");
    expect(inserts.length).toBeGreaterThanOrEqual(1);
    expect(inserts.some((c) => c.summary?.includes("Z99"))).toBe(true);
  });

  it("reports an updated cell as node-updated when overwriting", async () => {
    const a = await loadAgent("01-single-sheet-numbers.xlsx");
    const sheetName = a.listSheets()[0]!.name;
    await a.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet: sheetName, ref: "AA1", value: "first" },
    });
    const before = a.getSnapshot();
    await a.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet: sheetName, ref: "AA1", value: "second" },
    });
    const after = a.getSnapshot();
    const diff = diffXlsxSnapshots(before, after);
    const updates = diff.changes.filter((c) => c.kind === "node-updated");
    expect(updates).toHaveLength(1);
    expect(updates[0].summary).toMatch(/AA1/);
    expect(updates[0].summary).toMatch(/second/);
  });

  it("reports cleared cells as node-deleted", async () => {
    const a = await loadAgent("01-single-sheet-numbers.xlsx");
    const sheetName = a.listSheets()[0]!.name;
    await a.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet: sheetName, ref: "AA1", value: 99 },
    });
    const before = a.getSnapshot();
    await a.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet: sheetName, ref: "AA1", value: null },
    });
    const after = a.getSnapshot();
    const diff = diffXlsxSnapshots(before, after);
    const deletes = diff.changes.filter((c) => c.kind === "node-deleted");
    expect(deletes).toHaveLength(1);
    expect(deletes[0].summary).toMatch(/AA1/);
  });

  it("reports a sheet rename as node-updated on name", async () => {
    const a = await loadAgent("02-multi-sheet.xlsx");
    const sheetName = a.listSheets()[0]!.name;
    const before = a.getSnapshot();
    await a.applyCommand({
      type: "xlsx:rename-sheet",
      payload: { name: sheetName, newName: "Renamed!" },
    });
    const after = a.getSnapshot();
    const diff = diffXlsxSnapshots(before, after);
    const updates = diff.changes.filter((c) => c.kind === "node-updated");
    expect(updates.some((c) => c.summary?.includes("Renamed!"))).toBe(true);
  });

  it("reports merge add and remove", async () => {
    const a = await loadAgent("01-single-sheet-numbers.xlsx");
    const sheetName = a.listSheets()[0]!.name;
    const before = a.getSnapshot();
    await a.applyCommand({
      type: "xlsx:merge-cells",
      payload: { sheet: sheetName, range: "AC1:AD2" },
    });
    const afterMerge = a.getSnapshot();
    const diff1 = diffXlsxSnapshots(before, afterMerge);
    expect(diff1.changes.some((c) => c.kind === "node-inserted" && c.summary?.includes("merge"))).toBe(true);

    await a.applyCommand({
      type: "xlsx:unmerge-cells",
      payload: { sheet: sheetName, range: "AC1:AD2" },
    });
    const afterUnmerge = a.getSnapshot();
    const diff2 = diffXlsxSnapshots(afterMerge, afterUnmerge);
    expect(diff2.changes.some((c) => c.kind === "node-deleted" && c.summary?.includes("merge"))).toBe(true);
  });

  it("returns no changes when snapshots are identical", async () => {
    const a = await loadAgent("01-single-sheet-numbers.xlsx");
    const snap = a.getSnapshot();
    const diff = diffXlsxSnapshots(snap, snap);
    expect(diff.changes).toHaveLength(0);
  });
});
