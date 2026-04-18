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

describe("xlsx:insert-row — happy path", () => {
  it("shifts every cell at or below the insertion point down by `count`", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "Z5", value: "marker" },
    });
    const m = await bus.dispatch({
      type: "xlsx:insert-row",
      payload: { sheet: "Inventory", at: 3, count: 2 },
    });
    expect(m.status).toBe("approved");
    const snap = bus.getWorking();
    // Z5 (row 4) → Z7 (row 6).
    expect(cellAt(snap, "Inventory", 6, 25)?.value).toBe("marker");
    expect(cellAt(snap, "Inventory", 4, 25)).toBeUndefined();
  });

  it("leaves cells above the insertion point untouched", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "Z1", value: 11 },
    });
    await bus.dispatch({
      type: "xlsx:insert-row",
      payload: { sheet: "Inventory", at: 5, count: 1 },
    });
    expect(cellAt(bus.getWorking(), "Inventory", 0, 25)?.value).toBe(11);
  });

  it("rewrites a formula whose range spans the insertion point", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    for (let i = 1; i <= 10; i++) {
      await bus.dispatch({
        type: "xlsx:set-cell-value",
        payload: { sheet: "Inventory", ref: `Y${i}`, value: i },
      });
    }
    await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=Y1+SUM(Y2:Y10)" },
    });
    const m = await bus.dispatch({
      type: "xlsx:insert-row",
      payload: { sheet: "Inventory", at: 2, count: 1 },
    });
    expect(m.status).toBe("approved");
    // Z1 was at row 0 (above at0=1), so it stays at Z1.
    const z1 = cellAt(bus.getWorking(), "Inventory", 0, 25);
    expect(z1?.formula?.text).toBe("Y1+SUM(Y3:Y11)");
    // The total stays the same — Y1..Y10 just shifted to Y1, Y3..Y11.
    expect(z1?.value).toBe(55);
  });

  it("emits a rows-inserted summary in the diff", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:insert-row",
      payload: { sheet: "Inventory", at: 3, count: 2 },
    });
    expect(m.status).toBe("approved");
    const summary = m.diff?.changes.find((c) => c.field === "rows-inserted");
    expect(summary).toBeDefined();
    expect(summary?.meta).toMatchObject({ at: 3, count: 2, sheet: "Inventory" });
  });

  it("does not rewrite cross-sheet refs targeting other sheets", async () => {
    const { bus } = await makeBus("02-multi-sheet.xlsx");
    const initial = bus.getWorking();
    const sheets = initial.root.sheets;
    expect(sheets.length).toBeGreaterThanOrEqual(2);
    const otherName = sheets[1].name;
    const otherRef: string = `${otherName}!A1`;
    await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: sheets[0].name, ref: "Z1", formula: `=${otherRef}` },
    });
    const before = cellAt(bus.getWorking(), sheets[0].name, 0, 25)?.formula?.text;
    await bus.dispatch({
      type: "xlsx:insert-row",
      payload: { sheet: sheets[0].name, at: 1, count: 1 },
    });
    // Z1 shifts to Z2 because it's at the insertion point.
    const shifted = cellAt(bus.getWorking(), sheets[0].name, 1, 25);
    expect(shifted?.formula?.text).toBe(before);
  });
});

describe("xlsx:insert-row — merges", () => {
  it("expands a merge whose bottom edge equals the insertion point", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "Y3:Y5" },
    });
    const m = await bus.dispatch({
      type: "xlsx:insert-row",
      payload: { sheet: "Inventory", at: 5, count: 1 },
    });
    expect(m.status).toBe("approved");
    const sheet = bus.getWorking().root.sheets.find((s) => s.name === "Inventory")!;
    // at=5 → at0=4. Original merge r1=2,r2=4. After: r1=2, r2=5.
    const merge = sheet.merges.find((mm) => mm.r1 === 2);
    expect(merge).toBeDefined();
    expect(merge?.r2).toBe(5);
  });

  it("shifts a merge entirely below the insertion point", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "Y10:Y12" },
    });
    await bus.dispatch({
      type: "xlsx:insert-row",
      payload: { sheet: "Inventory", at: 5, count: 2 },
    });
    const sheet = bus.getWorking().root.sheets.find((s) => s.name === "Inventory")!;
    const merge = sheet.merges.find((mm) => mm.c1 === 24);
    expect(merge).toEqual({ r1: 11, c1: 24, r2: 13, c2: 24 });
  });

  it("rejects an insertion that would split a merge mid-region", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "Y3:Y8" },
    });
    const m = await bus.dispatch({
      type: "xlsx:insert-row",
      payload: { sheet: "Inventory", at: 5, count: 1 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("merge-boundary-crossed");
  });
});

describe("xlsx:insert-row — validation", () => {
  it("rejects an unknown sheet", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:insert-row",
      payload: { sheet: "Nope", at: 1, count: 1 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-sheet");
  });

  it("rejects at < 1", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:insert-row",
      payload: { sheet: "Inventory", at: 0, count: 1 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-position");
  });

  it("rejects count < 1", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:insert-row",
      payload: { sheet: "Inventory", at: 1, count: 0 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-count");
  });

  it("rejects at + count - 1 > 1048576", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:insert-row",
      payload: { sheet: "Inventory", at: 1048576, count: 2 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-count");
  });
});

describe("xlsx:insert-row — round trip", () => {
  it("preserves dependent recalc through the cell-update pipeline", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "Y2", value: 7 },
    });
    await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=Y2*3" },
    });
    expect(cellAt(bus.getWorking(), "Inventory", 0, 25)?.value as CellValue).toBe(21);
    await bus.dispatch({
      type: "xlsx:insert-row",
      payload: { sheet: "Inventory", at: 2, count: 1 },
    });
    // Y2 shifted to Y3; formula rewritten to =Y3*3; value still 21.
    const z1 = cellAt(bus.getWorking(), "Inventory", 0, 25);
    expect(z1?.formula?.text).toBe("Y3*3");
    expect(z1?.value).toBe(21);
  });
});
