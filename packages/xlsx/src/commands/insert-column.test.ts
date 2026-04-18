import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { CommandBus, defaultIdMinter } from "@officeai/core";
import { cellKey } from "../model/refs.js";
import type { XlsxSnapshot } from "../model/types.js";
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

describe("xlsx:insert-column — happy path", () => {
  it("shifts every cell at or right of the insertion point by `count`", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "Z1", value: "marker" },
    });
    const m = await bus.dispatch({
      type: "xlsx:insert-column",
      payload: { sheet: "Inventory", at: 26, count: 1 },
    });
    expect(m.status).toBe("approved");
    expect(cellAt(bus.getWorking(), "Inventory", 0, 26)?.value).toBe("marker");
    expect(cellAt(bus.getWorking(), "Inventory", 0, 25)).toBeUndefined();
  });

  it("leaves cells left of the insertion untouched", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "A1", value: "left" },
    });
    await bus.dispatch({
      type: "xlsx:insert-column",
      payload: { sheet: "Inventory", at: 5, count: 1 },
    });
    expect(cellAt(bus.getWorking(), "Inventory", 0, 0)?.value).toBe("left");
  });

  it("rewrites a formula whose range spans the insertion point", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    for (let i = 0; i < 5; i++) {
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
      type: "xlsx:insert-column",
      payload: { sheet: "Inventory", at: 22, count: 1 },
    });
    // U1 (col 20) sits left of the insertion at col 22 so its left edge
    // is unchanged; Y1 (col 24) shifts right to Z1 (col 25). Z1 itself
    // shifts to AA1 (col 26) where the formula now lives.
    const aa1 = cellAt(bus.getWorking(), "Inventory", 0, 26);
    expect(aa1?.formula?.text).toBe("SUM(U1:Z1)");
    expect(aa1?.value).toBe(15);
  });

  it("emits a columns-inserted summary in the diff", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:insert-column",
      payload: { sheet: "Inventory", at: 4, count: 3 },
    });
    expect(m.status).toBe("approved");
    const summary = m.diff?.changes.find((c) => c.field === "columns-inserted");
    expect(summary).toBeDefined();
    expect(summary?.meta).toMatchObject({ at: 4, count: 3, sheet: "Inventory" });
  });
});

describe("xlsx:insert-column — merges", () => {
  it("expands a merge whose right edge equals the insertion point", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "Y1:AA1" },
    });
    await bus.dispatch({
      type: "xlsx:insert-column",
      payload: { sheet: "Inventory", at: 27, count: 1 },
    });
    const sheet = bus.getWorking().root.sheets.find((s) => s.name === "Inventory")!;
    const merge = sheet.merges.find((mm) => mm.r1 === 0 && mm.c1 === 24);
    // Original c1=24,c2=26. at0=26 == c2 → expand. New c2=27.
    expect(merge?.c2).toBe(27);
  });

  it("rejects an insertion that would split a merge mid-region", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "Y1:AC1" },
    });
    const m = await bus.dispatch({
      type: "xlsx:insert-column",
      payload: { sheet: "Inventory", at: 27, count: 1 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("merge-boundary-crossed");
  });
});

describe("xlsx:insert-column — validation", () => {
  it("rejects an unknown sheet", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:insert-column",
      payload: { sheet: "Nope", at: 1, count: 1 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-sheet");
  });

  it("rejects at < 1 or count < 1", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m1 = await bus.dispatch({
      type: "xlsx:insert-column",
      payload: { sheet: "Inventory", at: 0, count: 1 },
    });
    expect(m1.rejection?.code).toBe("invalid-position");
    const m2 = await bus.dispatch({
      type: "xlsx:insert-column",
      payload: { sheet: "Inventory", at: 1, count: 0 },
    });
    expect(m2.rejection?.code).toBe("invalid-count");
  });

  it("rejects at + count - 1 > 16384", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:insert-column",
      payload: { sheet: "Inventory", at: 16384, count: 2 },
    });
    expect(m.rejection?.code).toBe("invalid-count");
  });
});
