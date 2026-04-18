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

describe("xlsx:delete-column — happy path", () => {
  it("drops cells inside the deletion band", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "Y1", value: "doomed" },
    });
    await bus.dispatch({
      type: "xlsx:delete-column",
      payload: { sheet: "Inventory", at: 25, count: 1 },
    });
    expect(cellAt(bus.getWorking(), "Inventory", 0, 24)).toBeUndefined();
  });

  it("shifts cells right of the deletion band left by `count`", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "Z1", value: "right" },
    });
    await bus.dispatch({
      type: "xlsx:delete-column",
      payload: { sheet: "Inventory", at: 25, count: 1 },
    });
    expect(cellAt(bus.getWorking(), "Inventory", 0, 24)?.value).toBe("right");
    expect(cellAt(bus.getWorking(), "Inventory", 0, 25)).toBeUndefined();
  });

  it("leaves cells left of the deletion band untouched", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "A1", value: "left" },
    });
    await bus.dispatch({
      type: "xlsx:delete-column",
      payload: { sheet: "Inventory", at: 5, count: 1 },
    });
    expect(cellAt(bus.getWorking(), "Inventory", 0, 0)?.value).toBe("left");
  });

  it("emits a columns-deleted summary in the diff", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:delete-column",
      payload: { sheet: "Inventory", at: 5, count: 2 },
    });
    expect(m.status).toBe("approved");
    const summary = m.diff?.changes.find((c) => c.field === "columns-deleted");
    expect(summary).toBeDefined();
    expect(summary?.meta).toMatchObject({ at: 5, count: 2, sheet: "Inventory" });
  });
});

describe("xlsx:delete-column — formula rewrite", () => {
  it("shifts a range that overlaps the deletion band partially", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    for (let i = 0; i < 5; i++) {
      // U1..Y1 = 1..5 (cols 20..24)
      await bus.dispatch({
        type: "xlsx:set-cell-value",
        payload: { sheet: "Inventory", ref: `${String.fromCharCode(85 + i)}1`, value: i + 1 },
      });
    }
    await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=SUM(U1:Y1)" },
    });
    await bus.dispatch({
      type: "xlsx:delete-column",
      payload: { sheet: "Inventory", at: 22, count: 2 },
    });
    // Z1 (col 25) shifts left to col 23 (X1).
    const x1 = cellAt(bus.getWorking(), "Inventory", 0, 23);
    expect(x1?.formula?.text).toBe("SUM(U1:W1)");
    // V,W (originally V=2, W=3) were deleted; X (=4) shifts to V; Y (=5)
    // shifts to W. New range U:W = 1 + 4 + 5 = 10.
    expect(x1?.value).toBe(10);
  });

  it("rewrites a deleted single-cell ref to #REF!", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "B1", value: 99 },
    });
    await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=B1+1" },
    });
    await bus.dispatch({
      type: "xlsx:delete-column",
      payload: { sheet: "Inventory", at: 2, count: 1 },
    });
    // Z1 was at col 25; deletion at col 1 (B) shifts Z1 to Y1 (col 24).
    const y1 = cellAt(bus.getWorking(), "Inventory", 0, 24);
    expect(y1?.formula?.text).toBe("#REF!+1");
    const value = y1?.value as CellValue;
    expect(typeof value === "object" && value && "kind" in value && value.code === "#REF!").toBe(true);
  });

  it("emits a referenced-cell-deleted change for casualties", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=B1" },
    });
    const m = await bus.dispatch({
      type: "xlsx:delete-column",
      payload: { sheet: "Inventory", at: 2, count: 1 },
    });
    expect(m.status).toBe("approved");
    const casualty = m.diff?.changes.find((c) => c.field === "referenced-cell-deleted");
    expect(casualty).toBeDefined();
  });

  it("does not rewrite cross-sheet refs targeting other sheets", async () => {
    const { bus } = await makeBus("02-multi-sheet.xlsx");
    const sheets = bus.getWorking().root.sheets;
    expect(sheets.length).toBeGreaterThanOrEqual(2);
    await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: sheets[1].name, ref: "Z1", formula: `=${sheets[0].name}!A1` },
    });
    await bus.dispatch({
      type: "xlsx:delete-column",
      payload: { sheet: sheets[1].name, at: 1, count: 1 },
    });
    // Z1 (col 25) shifts left to Y1 (col 24); cross-sheet ref unaffected.
    const y1 = cellAt(bus.getWorking(), sheets[1].name, 0, 24);
    expect(y1?.formula?.text).toBe(`${sheets[0].name}!A1`);
  });
});

describe("xlsx:delete-column — merges", () => {
  it("drops a merge fully inside the deletion band", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "Y1:AA1" },
    });
    await bus.dispatch({
      type: "xlsx:delete-column",
      payload: { sheet: "Inventory", at: 25, count: 3 },
    });
    const sheet = bus.getWorking().root.sheets.find((s) => s.name === "Inventory")!;
    expect(sheet.merges.find((m) => m.r1 === 0 && m.c1 === 24)).toBeUndefined();
  });

  it("shifts a merge entirely right of the deletion band left", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "Y1:AA1" },
    });
    await bus.dispatch({
      type: "xlsx:delete-column",
      payload: { sheet: "Inventory", at: 5, count: 2 },
    });
    const sheet = bus.getWorking().root.sheets.find((s) => s.name === "Inventory")!;
    const merge = sheet.merges.find((m) => m.r1 === 0);
    expect(merge).toEqual({ r1: 0, c1: 22, r2: 0, c2: 24 });
  });

  it("rejects a deletion that partially overlaps a merge", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "Y1:AC1" },
    });
    const m = await bus.dispatch({
      type: "xlsx:delete-column",
      payload: { sheet: "Inventory", at: 27, count: 2 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("merge-boundary-crossed");
  });
});

describe("xlsx:delete-column — validation", () => {
  it("rejects an unknown sheet", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:delete-column",
      payload: { sheet: "Nope", at: 1, count: 1 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-sheet");
  });

  it("rejects at < 1 or count < 1", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m1 = await bus.dispatch({
      type: "xlsx:delete-column",
      payload: { sheet: "Inventory", at: 0, count: 1 },
    });
    expect(m1.rejection?.code).toBe("invalid-position");
    const m2 = await bus.dispatch({
      type: "xlsx:delete-column",
      payload: { sheet: "Inventory", at: 1, count: 0 },
    });
    expect(m2.rejection?.code).toBe("invalid-count");
  });
});
