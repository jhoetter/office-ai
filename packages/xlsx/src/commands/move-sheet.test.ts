import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { CommandBus, defaultIdMinter } from "@officeai/core";
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

describe("xlsx:move-sheet", () => {
  it("reorders the named sheet and reindexes neighbours", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const before = initial.root.sheets.map((s) => s.name);
    expect(before.length).toBeGreaterThanOrEqual(2);
    const moving = before[0];

    const m = await bus.dispatch({
      type: "xlsx:move-sheet",
      payload: { name: moving, to: before.length - 1 },
    });
    expect(m.status).toBe("approved");

    const snap = bus.getWorking();
    const after = snap.root.sheets.map((s) => s.name);
    expect(after[after.length - 1]).toBe(moving);
    snap.root.sheets.forEach((s, i) => expect(s.index).toBe(i));
  });

  it("clamps out-of-range destinations to the workbook tail", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const target = initial.root.sheets[0];
    const m = await bus.dispatch({
      type: "xlsx:move-sheet",
      payload: { name: target.name, to: 9999 },
    });
    expect(m.status).toBe("approved");
    const snap = bus.getWorking();
    expect(snap.root.sheets[snap.root.sheets.length - 1].name).toBe(target.name);
  });

  it("is a no-op when source equals destination", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const target = initial.root.sheets[0];
    const m = await bus.dispatch({
      type: "xlsx:move-sheet",
      payload: { name: target.name, to: 0 },
    });
    expect(m.status).toBe("approved");
    expect(bus.getWorking().root.sheets[0].name).toBe(target.name);
  });

  it("rejects unknown sheet names", async () => {
    const { bus } = await makeBus("02-multi-sheet.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:move-sheet",
      payload: { name: "Phantom", to: 0 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-sheet");
  });

  it("round-trips through serialize/parse with the new order", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const before = initial.root.sheets.map((s) => s.name);
    const moving = before[0];
    const expected = [...before.slice(1), moving];

    await bus.dispatch({
      type: "xlsx:move-sheet",
      payload: { name: moving, to: before.length - 1 },
    });
    const snap = bus.getWorking();
    expect(snap.dirty.workbook).toBe(true);

    const out = await serializeXlsx(snap);
    const reparsed = await parseXlsx(new Uint8Array(out), { idMinter: defaultIdMinter });
    expect(reparsed.root.sheets.map((s) => s.name)).toEqual(expected);
  });
});
