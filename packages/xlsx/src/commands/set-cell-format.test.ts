import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { CommandBus, defaultIdMinter } from "@officeai/core";
import { cellKey } from "../model/refs.js";
import { flattenCellXf } from "../model/style-mutate.js";
import type { XlsxSnapshot } from "../model/types.js";
import { parseXlsx } from "../parser/index.js";
import { serializeXlsx } from "../serializer/index.js";
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

describe("xlsx:set-cell-format — bold", () => {
  it("bolds a single blank cell and grows cellXfs by one", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const before = initial.root.styles.cellXfs.length;
    const m = await bus.dispatch({
      type: "xlsx:set-cell-format",
      payload: {
        sheet: "Inventory",
        range: "Z1",
        format: { font: { bold: true } },
      },
    });
    expect(m.status).toBe("approved");
    const snap = bus.getWorking();
    const sheet = snap.root.sheets[0];
    const cell = sheet.cells.get(cellKey(0, 25));
    expect(cell).toBeDefined();
    expect(cell?.styleId).toBeDefined();
    expect(snap.root.styles.cellXfs.length).toBe(before + 1);
    const eff = flattenCellXf(snap.root.styles, cell?.styleId);
    expect(eff.font.bold).toBe(true);
    expect(snap.dirty.styles).toBe(true);
    expect(snap.dirty.sheets.has(sheet.partPath)).toBe(true);
  });

  it("dedupes when bolding an already-bold cell", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-format",
      payload: { sheet: "Inventory", range: "Z1", format: { font: { bold: true } } },
    });
    const after1 = bus.getWorking();
    const grew1 = after1.root.styles.cellXfs.length;
    const xf1 = after1.root.sheets[0].cells.get(cellKey(0, 25))?.styleId;

    const m = await bus.dispatch({
      type: "xlsx:set-cell-format",
      payload: { sheet: "Inventory", range: "Z1", format: { font: { bold: true } } },
    });
    expect(m.status).toBe("approved");
    const after2 = bus.getWorking();
    expect(after2.root.styles.cellXfs.length).toBe(grew1);
    expect(after2.root.sheets[0].cells.get(cellKey(0, 25))?.styleId).toBe(xf1);
    expect(m.diff?.changes).toHaveLength(0);
  });
});

describe("xlsx:set-cell-format — bulk fill across a range", () => {
  it("allocates a small constant number of new xfs over a uniform range", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const beforeCount = initial.root.styles.cellXfs.length;
    const m = await bus.dispatch({
      type: "xlsx:set-cell-format",
      payload: {
        sheet: "Inventory",
        range: "X1:X10",
        format: { fill: { color: "FFEEAA" } },
      },
    });
    expect(m.status).toBe("approved");
    const snap = bus.getWorking();
    const grew = snap.root.styles.cellXfs.length - beforeCount;
    expect(grew).toBe(1);
    for (let r = 0; r < 10; r++) {
      const cell = snap.root.sheets[0].cells.get(cellKey(r, 23));
      expect(cell?.styleId).toBeDefined();
      const eff = flattenCellXf(snap.root.styles, cell?.styleId);
      expect(eff.fill.fgColor?.rgb).toBe("FFFFEEAA");
    }
  });
});

describe("xlsx:set-cell-format — validation", () => {
  it("rejects a malformed colour with invalid-format", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-format",
      payload: {
        sheet: "Inventory",
        range: "Z1",
        format: { fill: { color: "not-hex" } },
      },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-format");
  });

  it("rejects an unknown sheet", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-format",
      payload: { sheet: "Nope", range: "A1", format: { font: { bold: true } } },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-sheet");
  });

  it("rejects an invalid range", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-format",
      payload: { sheet: "Inventory", range: "🚫", format: { font: { bold: true } } },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-range");
  });

  it("accepts a built-in numFmtId integer string", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-format",
      payload: { sheet: "Inventory", range: "Z1", format: { numberFormat: "9" } },
    });
    expect(m.status).toBe("approved");
    const snap = bus.getWorking();
    const cell = snap.root.sheets[0].cells.get(cellKey(0, 25));
    const eff = flattenCellXf(snap.root.styles, cell?.styleId);
    expect(eff.numFmtId).toBe(9);
  });

  it("registers a custom number format starting at 164", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-format",
      payload: {
        sheet: "Inventory",
        range: "Z1",
        format: { numberFormat: "0.0000" },
      },
    });
    expect(m.status).toBe("approved");
    const snap = bus.getWorking();
    const cell = snap.root.sheets[0].cells.get(cellKey(0, 25));
    const eff = flattenCellXf(snap.root.styles, cell?.styleId);
    expect(eff.numFmtId).toBeGreaterThanOrEqual(164);
    expect(snap.root.styles.numFmts.get(eff.numFmtId)).toBe("0.0000");
  });
});

describe("xlsx:set-cell-format — round trip", () => {
  it("survives parse → format → serialize → reparse with bold preserved", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "Z1", value: "header" },
    });
    await bus.dispatch({
      type: "xlsx:set-cell-format",
      payload: {
        sheet: "Inventory",
        range: "Z1",
        format: { font: { bold: true }, fill: { color: "FFEEAA" } },
      },
    });
    const out = await serializeXlsx(bus.getWorking());
    const reparsed = await parseXlsx(new Uint8Array(out));
    const cell = reparsed.root.sheets[0].cells.get(cellKey(0, 25));
    expect(cell?.styleId).toBeDefined();
    const eff = flattenCellXf(reparsed.root.styles, cell?.styleId);
    expect(eff.font.bold).toBe(true);
    expect(eff.fill.fgColor?.rgb).toBe("FFFFEEAA");
  });
});
