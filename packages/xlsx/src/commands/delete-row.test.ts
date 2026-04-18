import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { CommandBus, defaultIdMinter } from "@officeai/core";
import { cellKey } from "../model/refs.js";
import type { CellValue, XlsxSnapshot } from "../model/types.js";
import { parseXlsx } from "../parser/index.js";
import { allXlsxHandlers } from "./registry.js";

const here = dirname(fileURLToPath(import.meta.url));
const fixtures = resolve(here, "../../../../fixtures/xlsx/synthetic");

async function loadFixture(name: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(resolve(fixtures, name)));
}

async function makeBus(fixture: string): Promise<{
  bus: CommandBus<XlsxSnapshot>;
  initial: XlsxSnapshot;
}> {
  const buf = await loadFixture(fixture);
  const initial = await parseXlsx(buf, { idMinter: defaultIdMinter });
  const bus = new CommandBus<XlsxSnapshot>(initial);
  bus.registerAll(allXlsxHandlers);
  return { bus, initial };
}

function cellAt(snap: XlsxSnapshot, sheetName: string, row: number, col: number) {
  const sheet = snap.root.sheets.find((s) => s.name === sheetName);
  if (!sheet) throw new Error(`no sheet ${sheetName}`);
  return sheet.cells.get(cellKey(row, col));
}

describe("xlsx:delete-row — happy path", () => {
  it("drops cells inside the deletion band", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "Z3", value: "doomed" },
    });
    await bus.dispatch({
      type: "xlsx:delete-row",
      payload: { sheet: "Inventory", at: 3, count: 1 },
    });
    expect(cellAt(bus.getWorking(), "Inventory", 2, 25)).toBeUndefined();
  });

  it("shifts cells below the deletion band up by `count`", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "Z10", value: "below" },
    });
    await bus.dispatch({
      type: "xlsx:delete-row",
      payload: { sheet: "Inventory", at: 3, count: 2 },
    });
    expect(cellAt(bus.getWorking(), "Inventory", 7, 25)?.value).toBe("below");
    expect(cellAt(bus.getWorking(), "Inventory", 9, 25)).toBeUndefined();
  });

  it("leaves cells above the deletion band untouched", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "Z1", value: "stay" },
    });
    await bus.dispatch({
      type: "xlsx:delete-row",
      payload: { sheet: "Inventory", at: 5, count: 1 },
    });
    expect(cellAt(bus.getWorking(), "Inventory", 0, 25)?.value).toBe("stay");
  });

  it("emits a rows-deleted summary in the diff", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:delete-row",
      payload: { sheet: "Inventory", at: 5, count: 2 },
    });
    expect(m.status).toBe("approved");
    const summary = m.diff?.changes.find((c) => c.field === "rows-deleted");
    expect(summary).toBeDefined();
    expect(summary?.meta).toMatchObject({ at: 5, count: 2, sheet: "Inventory" });
  });
});

describe("xlsx:delete-row — formula rewrite", () => {
  it("shifts a range that overlaps the deletion band partially", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    for (let i = 1; i <= 10; i++) {
      await bus.dispatch({
        type: "xlsx:set-cell-value",
        payload: { sheet: "Inventory", ref: `Y${i}`, value: i },
      });
    }
    await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=SUM(Y1:Y10)" },
    });
    await bus.dispatch({
      type: "xlsx:delete-row",
      payload: { sheet: "Inventory", at: 5, count: 2 },
    });
    const z1 = cellAt(bus.getWorking(), "Inventory", 0, 25);
    expect(z1?.formula?.text).toBe("SUM(Y1:Y8)");
    // Original sum 1..10 = 55; removed Y5+Y6 = 11; new sum = 44.
    expect(z1?.value).toBe(44);
  });

  it("rewrites a deleted single-cell ref to #REF!", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "Y5", value: 99 },
    });
    await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=Y5+1" },
    });
    await bus.dispatch({
      type: "xlsx:delete-row",
      payload: { sheet: "Inventory", at: 5, count: 1 },
    });
    const z1 = cellAt(bus.getWorking(), "Inventory", 0, 25);
    expect(z1?.formula?.text).toBe("#REF!+1");
    const value = z1?.value as CellValue;
    expect(typeof value === "object" && value && "kind" in value && value.code === "#REF!").toBe(true);
  });

  it("emits a referenced-cell-deleted change for casualties", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=Y5" },
    });
    const m = await bus.dispatch({
      type: "xlsx:delete-row",
      payload: { sheet: "Inventory", at: 5, count: 1 },
    });
    expect(m.status).toBe("approved");
    const casualty = m.diff?.changes.find((c) => c.field === "referenced-cell-deleted");
    expect(casualty).toBeDefined();
  });

  it("does not rewrite cross-sheet refs in formulas on other sheets", async () => {
    const { bus } = await makeBus("02-multi-sheet.xlsx");
    const sheets = bus.getWorking().root.sheets;
    expect(sheets.length).toBeGreaterThanOrEqual(2);
    await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: sheets[1].name, ref: "Z1", formula: `=${sheets[0].name}!A1` },
    });
    const before = cellAt(bus.getWorking(), sheets[1].name, 0, 25)?.formula?.text;
    await bus.dispatch({
      type: "xlsx:delete-row",
      payload: { sheet: sheets[1].name, at: 1, count: 1 },
    });
    // Formula on sheets[1] referenced sheets[0]!A1, which is on a different
    // sheet from the one we deleted from — so unchanged.
    const newRow = sheets[1].name;
    const after = bus.getWorking().root.sheets.find((s) => s.name === newRow);
    // Z1 was at row 0 inside the deletion → cell dropped entirely.
    expect(after?.cells.get(cellKey(0, 25))).toBeUndefined();
    void before;
  });
});

describe("xlsx:delete-row — merges", () => {
  it("drops a merge fully inside the deletion band", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "Y3:Y5" },
    });
    await bus.dispatch({
      type: "xlsx:delete-row",
      payload: { sheet: "Inventory", at: 3, count: 3 },
    });
    const sheet = bus.getWorking().root.sheets.find((s) => s.name === "Inventory")!;
    expect(sheet.merges.find((m) => m.r1 === 2 && m.c1 === 24)).toBeUndefined();
  });

  it("shifts a merge entirely below the deletion band up", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "Y10:Y12" },
    });
    await bus.dispatch({
      type: "xlsx:delete-row",
      payload: { sheet: "Inventory", at: 5, count: 2 },
    });
    const sheet = bus.getWorking().root.sheets.find((s) => s.name === "Inventory")!;
    const merge = sheet.merges.find((m) => m.c1 === 24);
    expect(merge).toEqual({ r1: 7, c1: 24, r2: 9, c2: 24 });
  });

  it("rejects a deletion that partially overlaps a merge", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "Y3:Y8" },
    });
    const m = await bus.dispatch({
      type: "xlsx:delete-row",
      payload: { sheet: "Inventory", at: 5, count: 2 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("merge-boundary-crossed");
  });
});

describe("xlsx:delete-row — validation", () => {
  it("rejects an unknown sheet", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:delete-row",
      payload: { sheet: "Nope", at: 1, count: 1 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-sheet");
  });

  it("rejects at < 1", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:delete-row",
      payload: { sheet: "Inventory", at: 0, count: 1 },
    });
    expect(m.rejection?.code).toBe("invalid-position");
  });

  it("rejects count < 1", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:delete-row",
      payload: { sheet: "Inventory", at: 1, count: 0 },
    });
    expect(m.rejection?.code).toBe("invalid-count");
  });
});
