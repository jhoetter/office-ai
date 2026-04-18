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

async function makeBus(): Promise<CommandBus<XlsxSnapshot>> {
  const buf = await readFile(resolve(fixtures, "01-single-sheet-numbers.xlsx"));
  const initial = await parseXlsx(new Uint8Array(buf), { idMinter: defaultIdMinter });
  const bus = new CommandBus<XlsxSnapshot>(initial);
  bus.registerAll(allXlsxHandlers);
  return bus;
}

async function seedRange(
  bus: CommandBus<XlsxSnapshot>,
  startRow: number,
  startCol: string,
  values: ReadonlyArray<unknown>,
  axis: "down" | "right" = "down"
): Promise<void> {
  for (let i = 0; i < values.length; i++) {
    const ref =
      axis === "down"
        ? `${startCol}${startRow + i}`
        : `${String.fromCharCode(startCol.charCodeAt(0) + i)}${startRow}`;
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref, value: values[i] as never },
    });
  }
}

describe("xlsx:fill-range — numeric series down", () => {
  it("extends an arithmetic progression", async () => {
    const bus = await makeBus();
    await seedRange(bus, 1, "Y", [10, 20]);
    const m = await bus.dispatch({
      type: "xlsx:fill-range",
      payload: { sheet: "Inventory", source: "Y1:Y2", target: "Y1:Y5", direction: "down" },
    });
    expect(m.status).toBe("approved");
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(2, 24))?.value).toBe(30);
    expect(sheet.cells.get(cellKey(3, 24))?.value).toBe(40);
    expect(sheet.cells.get(cellKey(4, 24))?.value).toBe(50);
  });

  it("repeats a single numeric sample", async () => {
    const bus = await makeBus();
    await seedRange(bus, 1, "Y", [7]);
    await bus.dispatch({
      type: "xlsx:fill-range",
      payload: { sheet: "Inventory", source: "Y1", target: "Y1:Y3", direction: "down" },
    });
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(1, 24))?.value).toBe(7);
    expect(sheet.cells.get(cellKey(2, 24))?.value).toBe(7);
  });
});

describe("xlsx:fill-range — weekday/month/text-numeric", () => {
  it("extends weekdays preserving casing", async () => {
    const bus = await makeBus();
    await seedRange(bus, 1, "Y", ["Mon", "Tue"]);
    await bus.dispatch({
      type: "xlsx:fill-range",
      payload: { sheet: "Inventory", source: "Y1:Y2", target: "Y1:Y5", direction: "down" },
    });
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(2, 24))?.value).toBe("Wed");
    expect(sheet.cells.get(cellKey(4, 24))?.value).toBe("Fri");
  });

  it("extends months", async () => {
    const bus = await makeBus();
    await seedRange(bus, 1, "Y", ["January", "February"]);
    await bus.dispatch({
      type: "xlsx:fill-range",
      payload: { sheet: "Inventory", source: "Y1:Y2", target: "Y1:Y4", direction: "down" },
    });
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(2, 24))?.value).toBe("March");
    expect(sheet.cells.get(cellKey(3, 24))?.value).toBe("April");
  });

  it("extends 'Item N' style strings", async () => {
    const bus = await makeBus();
    await seedRange(bus, 1, "Y", ["Item 1", "Item 2"]);
    await bus.dispatch({
      type: "xlsx:fill-range",
      payload: { sheet: "Inventory", source: "Y1:Y2", target: "Y1:Y4", direction: "down" },
    });
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(2, 24))?.value).toBe("Item 3");
    expect(sheet.cells.get(cellKey(3, 24))?.value).toBe("Item 4");
  });
});

describe("xlsx:fill-range — fill right", () => {
  it("extends a numeric series across columns", async () => {
    const bus = await makeBus();
    await seedRange(bus, 1, "Y", [1, 2], "right"); // Y1=1, Z1=2
    await bus.dispatch({
      type: "xlsx:fill-range",
      payload: { sheet: "Inventory", source: "Y1:Z1", target: "Y1:AB1", direction: "right" },
    });
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(0, 26))?.value).toBe(3); // AA1
    expect(sheet.cells.get(cellKey(0, 27))?.value).toBe(4); // AB1
  });
});

describe("xlsx:fill-range — fill up", () => {
  it("extends backwards (subtracts step)", async () => {
    const bus = await makeBus();
    await seedRange(bus, 5, "Y", [50, 60]); // Y5=50, Y6=60 (step 10)
    await bus.dispatch({
      type: "xlsx:fill-range",
      payload: { sheet: "Inventory", source: "Y5:Y6", target: "Y2:Y6", direction: "up" },
    });
    const sheet = bus.getWorking().root.sheets[0];
    // Filling up reverses samples, so step seen is -10; offsets walk
    // upward. The cells produced from Y4..Y2 should be 40, 30, 20.
    expect(sheet.cells.get(cellKey(3, 24))?.value).toBe(40);
    expect(sheet.cells.get(cellKey(2, 24))?.value).toBe(30);
    expect(sheet.cells.get(cellKey(1, 24))?.value).toBe(20);
  });
});

describe("xlsx:fill-range — formulas", () => {
  it("shifts relative refs when filling formulas down", async () => {
    const bus = await makeBus();
    await seedRange(bus, 1, "Y", [10, 20, 30]);
    await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "Z1", formula: "=Y1*2" },
    });
    await bus.dispatch({
      type: "xlsx:fill-range",
      payload: { sheet: "Inventory", source: "Z1", target: "Z1:Z3", direction: "down" },
    });
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(1, 25))?.formula?.text).toBe("Y2*2");
    expect(sheet.cells.get(cellKey(1, 25))?.value).toBe(40);
    expect(sheet.cells.get(cellKey(2, 25))?.value).toBe(60);
  });
});

describe("xlsx:fill-range — validation", () => {
  it("rejects when target does not enclose source", async () => {
    const bus = await makeBus();
    await seedRange(bus, 1, "Y", [1, 2]);
    const m = await bus.dispatch({
      type: "xlsx:fill-range",
      payload: { sheet: "Inventory", source: "Y1:Y2", target: "Z1:Z3", direction: "down" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-range");
  });

  it("rejects mismatched direction vs geometry", async () => {
    const bus = await makeBus();
    await seedRange(bus, 1, "Y", [1, 2]);
    const m = await bus.dispatch({
      type: "xlsx:fill-range",
      payload: { sheet: "Inventory", source: "Y1:Y2", target: "Y1:Y4", direction: "right" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-range");
  });
});
