import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { CommandBus, defaultIdMinter } from "@officeai/core";
import { parseXlsx } from "../parser/index.js";
import { cellKey } from "../model/refs.js";
import type { CellValue, XlsxSnapshot } from "../model/types.js";
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

describe("xlsx:set-cell-formula — basic write + recalc", () => {
  it("writes a literal arithmetic formula and stores the cached value", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=1+2*3" },
    });
    expect(m.status).toBe("approved");
    const snap = bus.getWorking();
    const cell = cellAt(snap, "Inventory", 0, 25);
    expect(cell?.formula?.text).toBe("1+2*3");
    expect(cell?.value).toBe(7);
    expect(snap.dirty.sheets.has("xl/worksheets/sheet1.xml")).toBe(true);
  });

  it("accepts a formula without a leading =", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "10/4" },
    });
    expect(m.status).toBe("approved");
    const cell = cellAt(bus.getWorking(), "Inventory", 0, 25);
    expect(cell?.formula?.text).toBe("10/4");
    expect(cell?.value).toBe(2.5);
  });

  it("preserves the formula text verbatim (no whitespace canonicalisation)", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "= 1   +   2" },
    });
    const cell = cellAt(bus.getWorking(), "Inventory", 0, 25);
    expect(cell?.formula?.text).toBe(" 1   +   2");
    expect(cell?.value).toBe(3);
  });
});

describe("xlsx:set-cell-formula — references + recalc", () => {
  it("evaluates a SUM over a referenced range", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    // Seed a few literal numbers, then sum them.
    for (let i = 1; i <= 3; i++) {
      await bus.dispatch({
        type: "xlsx:set-cell-value",
        payload: { sheet: "Inventory", ref: `Y${i}`, value: i * 10 },
      });
    }
    const m = await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=SUM(Y1:Y3)" },
    });
    expect(m.status).toBe("approved");
    const cell = cellAt(bus.getWorking(), "Inventory", 0, 25);
    expect(cell?.value).toBe(60);
  });

  it("propagates downstream cached values when a formula's dependencies change later", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "Y1", value: 5 },
    });
    await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=Y1*2" },
    });
    expect(cellAt(bus.getWorking(), "Inventory", 0, 25)?.value).toBe(10);

    // Now author a *new* formula whose evaluation should re-fire the
    // existing dependent (Z1) because we recalc the whole graph.
    const m = await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z2", formula: "=Z1+100" },
    });
    expect(m.status).toBe("approved");
    expect(cellAt(bus.getWorking(), "Inventory", 1, 25)?.value).toBe(110);
  });
});

describe("xlsx:set-cell-formula — validation + errors", () => {
  it("rejects a malformed formula with formula-parse-error", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=*1" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("formula-parse-error");
  });

  it("rejects an unknown sheet", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Nope", ref: "A1", formula: "=1" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-sheet");
  });

  it("rejects an invalid ref", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "🚫", formula: "=1" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-ref");
  });

  it("clears the cell when the formula body is empty", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "Z1", value: 99 },
    });
    expect(cellAt(bus.getWorking(), "Inventory", 0, 25)?.value).toBe(99);
    const m = await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=" },
    });
    expect(m.status).toBe("approved");
    expect(cellAt(bus.getWorking(), "Inventory", 0, 25)).toBeUndefined();
  });

  it("surfaces #DIV/0! as a CellErrorValue in the cached value", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=1/0" },
    });
    expect(m.status).toBe("approved");
    const v = cellAt(bus.getWorking(), "Inventory", 0, 25)?.value as CellValue;
    expect(typeof v === "object" && v && "kind" in v && v.kind === "error").toBe(true);
    if (typeof v === "object" && v && "kind" in v) {
      expect(v.code).toBe("#DIV/0!");
    }
  });
});

describe("xlsx:set-cell-formula — cycles", () => {
  it("flags a self-loop with a circular diff entry", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=Z1+1" },
    });
    expect(m.status).toBe("approved");
    const cell = cellAt(bus.getWorking(), "Inventory", 0, 25);
    expect(cell?.value).toMatchObject({ kind: "error", code: "#REF!" });
    const circular = m.diff?.changes.find((c) => c.field === "circular");
    expect(circular).toBeDefined();
    expect(circular?.meta?.cycle).toBeDefined();
  });
});
