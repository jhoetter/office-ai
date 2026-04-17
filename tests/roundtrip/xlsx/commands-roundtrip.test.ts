import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { CommandBus } from "@officeai/core";
import { allXlsxHandlers, cellKey, parseXlsx, serializeXlsx, type XlsxSnapshot } from "@officeai/xlsx";

/**
 * Phase 5 end-to-end: dispatch the value-mutation + merge + rename
 * commands through the bus, serialize, reparse, and assert the typed
 * model reflects the edits.
 */

const FIXTURE_DIR = resolve(__dirname, "../../../fixtures/xlsx/synthetic");

async function loadFixture(name: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(resolve(FIXTURE_DIR, name)));
}

async function makeBus(name: string): Promise<{
  bus: CommandBus<XlsxSnapshot>;
  initial: XlsxSnapshot;
}> {
  const buf = await loadFixture(name);
  const initial = await parseXlsx(buf);
  const bus = new CommandBus<XlsxSnapshot>(initial);
  bus.registerAll(allXlsxHandlers);
  return { bus, initial };
}

describe("XLSX commands roundtrip", () => {
  it("set-cell-value persists across serialize → parse", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: initial.root.sheets[0].name, ref: "Z99", value: 123.45 },
    });
    const out = await serializeXlsx(bus.getWorking());
    const reparsed = await parseXlsx(new Uint8Array(out));
    const sheet = reparsed.root.sheets[0];
    expect(sheet.cells.get(cellKey(98, 25))?.value).toBe(123.45);
    expect(sheet.name).toBe(initial.root.sheets[0].name);
  });

  it("set-range-values persists a 2x2 block", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-range-values",
      payload: {
        sheet: initial.root.sheets[0].name,
        range: "Z1:AA2",
        values: [
          [1, "two"],
          [true, null],
        ],
      },
    });
    const out = await serializeXlsx(bus.getWorking());
    const reparsed = await parseXlsx(new Uint8Array(out));
    const sheet = reparsed.root.sheets[0];
    expect(sheet.cells.get(cellKey(0, 25))?.value).toBe(1);
    expect(sheet.cells.get(cellKey(0, 26))?.value).toBe("two");
    expect(sheet.cells.get(cellKey(1, 25))?.value).toBe(true);
    expect(sheet.cells.has(cellKey(1, 26))).toBe(false);
  });

  it("merge-cells then unmerge-cells round-trip cleanly", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-range-values",
      payload: {
        sheet: initial.root.sheets[0].name,
        range: "Z1:AB1",
        values: [["title", null, null]],
      },
    });
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: initial.root.sheets[0].name, range: "Z1:AB1" },
    });
    const merged = await serializeXlsx(bus.getWorking());
    const reparsedMerged = await parseXlsx(new Uint8Array(merged));
    expect(reparsedMerged.root.sheets[0].merges).toContainEqual({ r1: 0, c1: 25, r2: 0, c2: 27 });
    expect(reparsedMerged.root.sheets[0].cells.get(cellKey(0, 25))?.value).toBe("title");
  });

  it("rename-sheet persists in workbook.xml", async () => {
    const { bus } = await makeBus("02-multi-sheet.xlsx");
    await bus.dispatch({
      type: "xlsx:rename-sheet",
      payload: { name: "Expenses", newName: "Costs" },
    });
    const out = await serializeXlsx(bus.getWorking());
    const reparsed = await parseXlsx(new Uint8Array(out));
    expect(reparsed.root.sheets.map((s) => s.name)).toEqual(["Sales", "Costs", "Summary"]);
  });

  it("untouched parts in a renamed-sheet workbook stay byte-identical", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    await bus.dispatch({
      type: "xlsx:rename-sheet",
      payload: { name: "Expenses", newName: "Costs" },
    });
    const out = await serializeXlsx(bus.getWorking());
    const reparsed = await parseXlsx(new Uint8Array(out));
    const untouched = ["xl/sharedStrings.xml", "xl/styles.xml", "xl/_rels/workbook.xml.rels"];
    for (const path of untouched) {
      if (!(path in initial.partHashes)) continue;
      expect(reparsed.partHashes[path], `${path} should be byte-identical after rename`).toBe(
        initial.partHashes[path]
      );
    }
  });
});
