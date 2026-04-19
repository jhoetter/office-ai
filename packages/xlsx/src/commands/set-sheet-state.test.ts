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

describe("xlsx:set-sheet-state", () => {
  it("hides a non-active sheet and emits the workbook dirty bit", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const target = initial.root.sheets[1];

    const m = await bus.dispatch({
      type: "xlsx:set-sheet-state",
      payload: { name: target.name, state: "hidden" },
    });
    expect(m.status).toBe("approved");

    const snap = bus.getWorking();
    const after = snap.root.sheets.find((s) => s.name === target.name);
    expect(after?.state).toBe("hidden");
    expect(snap.dirty.workbook).toBe(true);
  });

  it("rejects hiding the only visible sheet", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    // Hide every sheet except the first; the next hide must fail.
    const visible = initial.root.sheets.filter((s) => s.state === "visible");
    for (let i = 1; i < visible.length; i++) {
      await bus.dispatch({
        type: "xlsx:set-sheet-state",
        payload: { name: visible[i].name, state: "hidden" },
      });
    }
    const m = await bus.dispatch({
      type: "xlsx:set-sheet-state",
      payload: { name: visible[0].name, state: "hidden" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-position");
  });

  it("round-trips through serialize/parse with the hidden state preserved", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const target = initial.root.sheets[1];
    await bus.dispatch({
      type: "xlsx:set-sheet-state",
      payload: { name: target.name, state: "hidden" },
    });
    const out = await serializeXlsx(bus.getWorking());
    const reparsed = await parseXlsx(new Uint8Array(out), { idMinter: defaultIdMinter });
    const echo = reparsed.root.sheets.find((s) => s.name === target.name);
    expect(echo?.state).toBe("hidden");
  });
});
