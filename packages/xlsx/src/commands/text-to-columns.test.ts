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

async function seedColumn(
  bus: CommandBus<XlsxSnapshot>,
  values: ReadonlyArray<string>,
  col: string = "Z"
): Promise<void> {
  for (let i = 0; i < values.length; i++) {
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: `${col}${i + 1}`, value: values[i]! },
    });
  }
}

describe("xlsx:text-to-columns", () => {
  it("splits a single-row CSV-like string in place", async () => {
    const bus = await makeBus();
    await seedColumn(bus, ["a,b,c"]);
    const m = await bus.dispatch({
      type: "xlsx:text-to-columns",
      payload: { sheet: "Inventory", range: "Z1", delimiter: "," },
    });
    expect(m.status).toBe("approved");
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(0, 25))?.value).toBe("a");
    expect(sheet.cells.get(cellKey(0, 26))?.value).toBe("b");
    expect(sheet.cells.get(cellKey(0, 27))?.value).toBe("c");
  });

  it("coerces numeric and boolean literals back into typed values", async () => {
    const bus = await makeBus();
    await seedColumn(bus, ["1;2.5;true;hello"]);
    const m = await bus.dispatch({
      type: "xlsx:text-to-columns",
      payload: { sheet: "Inventory", range: "Z1", delimiter: ";" },
    });
    expect(m.status).toBe("approved");
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(0, 25))?.value).toBe(1);
    expect(sheet.cells.get(cellKey(0, 26))?.value).toBe(2.5);
    expect(sheet.cells.get(cellKey(0, 27))?.value).toBe(true);
    expect(sheet.cells.get(cellKey(0, 28))?.value).toBe("hello");
  });

  it("writes results to a separate destination without disturbing the source", async () => {
    const bus = await makeBus();
    await seedColumn(bus, ["x|y|z"], "A");
    const m = await bus.dispatch({
      type: "xlsx:text-to-columns",
      payload: { sheet: "Inventory", range: "A1", delimiter: "|", destination: "C5" },
    });
    expect(m.status).toBe("approved");
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(0, 0))?.value).toBe("x|y|z");
    expect(sheet.cells.get(cellKey(4, 2))?.value).toBe("x");
    expect(sheet.cells.get(cellKey(4, 3))?.value).toBe("y");
    expect(sheet.cells.get(cellKey(4, 4))?.value).toBe("z");
  });

  it("collapses consecutive delimiters when treatConsecutiveAsOne is true", async () => {
    const bus = await makeBus();
    await seedColumn(bus, ["a,,,b,,c"]);
    const m = await bus.dispatch({
      type: "xlsx:text-to-columns",
      payload: {
        sheet: "Inventory",
        range: "Z1",
        delimiter: ",",
        treatConsecutiveAsOne: true,
      },
    });
    expect(m.status).toBe("approved");
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(0, 25))?.value).toBe("a");
    expect(sheet.cells.get(cellKey(0, 26))?.value).toBe("b");
    expect(sheet.cells.get(cellKey(0, 27))?.value).toBe("c");
    expect(sheet.cells.get(cellKey(0, 28))).toBeUndefined();
  });

  it("processes a multi-row range and pads to the widest split", async () => {
    const bus = await makeBus();
    await seedColumn(bus, ["a,b", "x,y,z", "lonely"]);
    const m = await bus.dispatch({
      type: "xlsx:text-to-columns",
      payload: { sheet: "Inventory", range: "Z1:Z3", delimiter: "," },
    });
    expect(m.status).toBe("approved");
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(0, 25))?.value).toBe("a");
    expect(sheet.cells.get(cellKey(0, 26))?.value).toBe("b");
    expect(sheet.cells.get(cellKey(1, 25))?.value).toBe("x");
    expect(sheet.cells.get(cellKey(1, 26))?.value).toBe("y");
    expect(sheet.cells.get(cellKey(1, 27))?.value).toBe("z");
    expect(sheet.cells.get(cellKey(2, 25))?.value).toBe("lonely");
    expect(sheet.cells.get(cellKey(2, 26))).toBeUndefined();
  });

  it("rejects an empty delimiter", async () => {
    const bus = await makeBus();
    await seedColumn(bus, ["abc"]);
    const m = await bus.dispatch({
      type: "xlsx:text-to-columns",
      payload: { sheet: "Inventory", range: "Z1", delimiter: "" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("supports multi-character delimiters", async () => {
    const bus = await makeBus();
    await seedColumn(bus, ["one :: two :: three"]);
    const m = await bus.dispatch({
      type: "xlsx:text-to-columns",
      payload: { sheet: "Inventory", range: "Z1", delimiter: " :: " },
    });
    expect(m.status).toBe("approved");
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(0, 25))?.value).toBe("one");
    expect(sheet.cells.get(cellKey(0, 26))?.value).toBe("two");
    expect(sheet.cells.get(cellKey(0, 27))?.value).toBe("three");
  });
});
