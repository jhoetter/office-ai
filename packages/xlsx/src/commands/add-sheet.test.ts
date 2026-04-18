import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { CommandBus, defaultIdMinter } from "@officeai/core";
import { cellKey } from "../model/refs.js";
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

describe("xlsx:add-sheet — placement", () => {
  it("appends a new sheet when `at` is omitted", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const before = initial.root.sheets.length;
    const m = await bus.dispatch({
      type: "xlsx:add-sheet",
      payload: { name: "Forecast" },
    });
    expect(m.status).toBe("approved");
    const snap = bus.getWorking();
    expect(snap.root.sheets).toHaveLength(before + 1);
    const last = snap.root.sheets[snap.root.sheets.length - 1];
    expect(last.name).toBe("Forecast");
    expect(last.index).toBe(before);
    expect(last.cells.size).toBe(0);
    expect(last.merges).toHaveLength(0);
    expect(snap.dirty.workbook).toBe(true);
    expect(snap.dirty.rels).toBe(true);
    expect(snap.dirty.contentTypes).toBe(true);
    expect(snap.dirty.sheets.has(last.partPath)).toBe(true);
  });

  it("inserts at index 0 and shifts existing sheets right", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const beforeNames = initial.root.sheets.map((s) => s.name);
    const m = await bus.dispatch({
      type: "xlsx:add-sheet",
      payload: { name: "Cover", at: 0 },
    });
    expect(m.status).toBe("approved");
    const snap = bus.getWorking();
    expect(snap.root.sheets.map((s) => s.name)).toEqual(["Cover", ...beforeNames]);
    snap.root.sheets.forEach((s, i) => expect(s.index).toBe(i));
  });

  it("inserts at a middle index", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const beforeNames = initial.root.sheets.map((s) => s.name);
    const m = await bus.dispatch({
      type: "xlsx:add-sheet",
      payload: { name: "Middle", at: 1 },
    });
    expect(m.status).toBe("approved");
    const after = bus.getWorking().root.sheets.map((s) => s.name);
    expect(after).toEqual([beforeNames[0], "Middle", ...beforeNames.slice(1)]);
  });

  it("emits a node-inserted diff with name/at/sheetId/partPath", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:add-sheet",
      payload: { name: "Notes", at: 1 },
    });
    expect(m.status).toBe("approved");
    expect(m.diff.changes).toHaveLength(1);
    const change = m.diff.changes[0];
    expect(change.kind).toBe("node-inserted");
    expect(change.summary).toBe("Added sheet 'Notes' at index 1");
    expect(change.meta?.name).toBe("Notes");
    expect(change.meta?.at).toBe(1);
    expect(typeof change.meta?.sheetId).toBe("string");
    expect(String(change.meta?.partPath)).toMatch(/^xl\/worksheets\/sheet\d+\.xml$/);
  });
});

describe("xlsx:add-sheet — name + position validation", () => {
  it("rejects a name colliding with an existing sheet (case-insensitive)", async () => {
    const { bus } = await makeBus("02-multi-sheet.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:add-sheet",
      payload: { name: "expenses" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("duplicate-name");
  });

  it("rejects a name with forbidden characters", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:add-sheet",
      payload: { name: "Bad/Name" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-name");
  });

  it("rejects the reserved name 'History' (case-insensitive)", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:add-sheet",
      payload: { name: "history" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-name");
  });

  it("rejects an empty name", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:add-sheet",
      payload: { name: "" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-name");
  });

  it("rejects a name longer than 31 characters", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:add-sheet",
      payload: { name: "a".repeat(32) },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-name");
  });

  it("rejects a position below zero", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:add-sheet",
      payload: { name: "X", at: -1 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-position");
  });

  it("rejects a position past the end", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:add-sheet",
      payload: { name: "X", at: initial.root.sheets.length + 1 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-position");
  });
});

describe("xlsx:add-sheet — round-trip", () => {
  it("survives parse → add → serialize → re-parse with all sheets intact", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const originalNames = initial.root.sheets.map((s) => s.name);
    const originalCellCounts = initial.root.sheets.map((s) => s.cells.size);

    const m = await bus.dispatch({
      type: "xlsx:add-sheet",
      payload: { name: "NewTab", at: 1 },
    });
    expect(m.status).toBe("approved");

    const snap = bus.getWorking();
    const out = await serializeXlsx(snap);
    const reparsed = await parseXlsx(new Uint8Array(out));

    const expectedOrder = [originalNames[0], "NewTab", ...originalNames.slice(1)];
    expect(reparsed.root.sheets.map((s) => s.name)).toEqual(expectedOrder);

    const newSheet = reparsed.root.sheets.find((s) => s.name === "NewTab")!;
    expect(newSheet.cells.size).toBe(0);
    expect(newSheet.merges).toHaveLength(0);
    expect(newSheet.kind).toBe("worksheet");

    for (const original of initial.root.sheets) {
      const after = reparsed.root.sheets.find((s) => s.name === original.name)!;
      expect(after).toBeDefined();
      expect(after.cells.size).toBe(originalCellCounts[original.index]);
      for (const [key, cell] of original.cells) {
        const reCell = after.cells.get(key);
        expect(reCell, `missing cell ${original.name}!${key}`).toBeDefined();
        expect(reCell?.value).toEqual(cell.value);
      }
    }
  });

  it("appends to a single-sheet workbook and round-trips both sheets", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:add-sheet",
      payload: { name: "Scratch" },
    });
    expect(m.status).toBe("approved");
    const out = await serializeXlsx(bus.getWorking());
    const reparsed = await parseXlsx(new Uint8Array(out));
    const names = reparsed.root.sheets.map((s) => s.name);
    expect(names).toContain("Inventory");
    expect(names).toContain("Scratch");
    expect(names[names.length - 1]).toBe("Scratch");
    const scratch = reparsed.root.sheets.find((s) => s.name === "Scratch")!;
    expect(scratch.cells.size).toBe(0);
  });

  it("supports follow-up edits on the freshly added sheet", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const add = await bus.dispatch({
      type: "xlsx:add-sheet",
      payload: { name: "Targets" },
    });
    expect(add.status).toBe("approved");
    const write = await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Targets", ref: "B2", value: 42 },
    });
    expect(write.status).toBe("approved");
    const out = await serializeXlsx(bus.getWorking());
    const reparsed = await parseXlsx(new Uint8Array(out));
    const targets = reparsed.root.sheets.find((s) => s.name === "Targets")!;
    expect(targets.cells.get(cellKey(1, 1))?.value).toBe(42);
  });
});
