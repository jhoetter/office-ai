import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { CommandBus, CommandError, defaultIdMinter } from "@officeai/core";
import { parseXlsx } from "../parser/index.js";
import type { XlsxSnapshot } from "../model/types.js";
import { cellKey } from "../model/refs.js";
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

describe("xlsx:set-cell-value", () => {
  it("writes a number to a blank cell and bumps revision", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "Z99", value: 42 },
    });
    expect(m.status).toBe("approved");
    const snap = bus.getWorking();
    expect(snap.revision).toBe(1);
    const sheet = snap.root.sheets[0];
    expect(sheet.cells.get(cellKey(98, 25))?.value).toBe(42);
    expect(snap.dirty.sheets.has("xl/worksheets/sheet1.xml")).toBe(true);
  });

  it("overwrites an existing cell, preserving formulas dropped", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheet = initial.root.sheets[0];
    const targetKey = [...sheet.cells.keys()][0];
    expect(targetKey).toBeDefined();
    const before = sheet.cells.get(targetKey)!;
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: {
        sheet: "Inventory",
        ref: refOf(before.row, before.col),
        value: 999,
      },
    });
    const after = bus.getWorking().root.sheets[0].cells.get(targetKey)!;
    expect(after.value).toBe(999);
    expect(after.formula).toBeUndefined();
  });

  it("clears a cell when value is null", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheet = initial.root.sheets[0];
    const k = [...sheet.cells.keys()][0];
    const before = sheet.cells.get(k)!;
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: refOf(before.row, before.col), value: null },
    });
    expect(bus.getWorking().root.sheets[0].cells.has(k)).toBe(false);
  });

  it("rejects unknown sheet", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Nope", ref: "A1", value: 1 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-sheet");
  });

  it("rejects invalid A1 ref", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "A1:B2", value: 1 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-ref");
  });

  it("recomputes dependent formulas when an input value changes", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    // Seed: B2 := 42, E8 := =B2 (cached value 42 by recalc inside set-cell-formula)
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "B2", value: 42 },
    });
    await bus.dispatch({
      type: "xlsx:set-cell-formula",
      payload: { sheet: "Inventory", ref: "E8", formula: "=B2" },
    });
    const seeded = bus.getWorking().root.sheets[0];
    expect(seeded.cells.get(cellKey(7, 4))?.value).toBe(42);

    // Now bump B2 → 44 via set-cell-value. E8 must follow.
    const m = await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "B2", value: 44 },
    });
    expect(m.status).toBe("approved");
    const after = bus.getWorking().root.sheets[0];
    expect(after.cells.get(cellKey(1, 1))?.value).toBe(44);
    expect(after.cells.get(cellKey(7, 4))?.value).toBe(44);
    // The dependent's formula text must be preserved.
    expect(after.cells.get(cellKey(7, 4))?.formula?.text).toBe("B2");
  });

  it("rejects formula-shaped string values", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "A1", value: "=SUM(A1:A2)" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("formula-string");
  });
});

describe("xlsx:set-range-values", () => {
  it("writes a 2x3 block of values", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-range-values",
      payload: {
        sheet: "Inventory",
        range: "Z1:AB2",
        values: [
          [1, 2, 3],
          ["a", "b", "c"],
        ],
      },
    });
    expect(m.status).toBe("approved");
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(0, 25))?.value).toBe(1);
    expect(sheet.cells.get(cellKey(0, 27))?.value).toBe(3);
    expect(sheet.cells.get(cellKey(1, 25))?.value).toBe("a");
  });

  it("rejects on dimension mismatch", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-range-values",
      payload: {
        sheet: "Inventory",
        range: "A1:B2",
        values: [[1, 2]],
      },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("dimension-mismatch");
  });

  it("clears cells when a value is null", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheet = initial.root.sheets[0];
    const cellEntries = [...sheet.cells.entries()];
    const targets = cellEntries.slice(0, 2);
    expect(targets.length).toBe(2);
    const refStart = refOf(targets[0][1].row, targets[0][1].col);
    const refEnd = refOf(targets[1][1].row, targets[1][1].col);
    if (targets[0][1].row !== targets[1][1].row) return;
    await bus.dispatch({
      type: "xlsx:set-range-values",
      payload: {
        sheet: "Inventory",
        range: `${refStart}:${refEnd}`,
        values: [[null, null]],
      },
    });
    const next = bus.getWorking().root.sheets[0];
    expect(next.cells.has(targets[0][0])).toBe(false);
    expect(next.cells.has(targets[1][0])).toBe(false);
  });
});

describe("xlsx:merge-cells", () => {
  it("merges a 2x2 region and clears non-anchor cells", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-range-values",
      payload: {
        sheet: "Inventory",
        range: "Z1:AA2",
        values: [
          ["anchor", "x"],
          ["y", "z"],
        ],
      },
    });
    const m = await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "Z1:AA2" },
    });
    expect(m.status).toBe("approved");
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.merges).toContainEqual({ r1: 0, c1: 25, r2: 1, c2: 26 });
    expect(sheet.cells.get(cellKey(0, 25))?.value).toBe("anchor");
    expect(sheet.cells.has(cellKey(0, 26))).toBe(false);
    expect(sheet.cells.has(cellKey(1, 25))).toBe(false);
    expect(sheet.cells.has(cellKey(1, 26))).toBe(false);
  });

  it("rejects single-cell range", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "A1" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-range");
  });

  it("rejects overlap with existing merge", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "Z1:AA2" },
    });
    const m = await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "AA2:AB3" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("overlap-with-existing-merge");
  });

  it("set-cell-value on non-anchor merged cell is rejected", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "Z1:AA2" },
    });
    const m = await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "AA1", value: 1 },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("merged-non-anchor");
  });
});

describe("xlsx:unmerge-cells", () => {
  it("removes an exact-match merge", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "Z1:AA2" },
    });
    const m = await bus.dispatch({
      type: "xlsx:unmerge-cells",
      payload: { sheet: "Inventory", range: "Z1:AA2" },
    });
    expect(m.status).toBe("approved");
    expect(bus.getWorking().root.sheets[0].merges).toEqual([]);
  });

  it("rejects when no exact-match merge exists", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:unmerge-cells",
      payload: { sheet: "Inventory", range: "Z1:AA2" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("merge-not-found");
  });
});

describe("xlsx:rename-sheet", () => {
  it("renames a sheet and marks workbook dirty", async () => {
    const { bus } = await makeBus("02-multi-sheet.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:rename-sheet",
      payload: { name: "Expenses", newName: "Costs" },
    });
    expect(m.status).toBe("approved");
    const next = bus.getWorking();
    expect(next.root.sheets.map((s) => s.name)).toContain("Costs");
    expect(next.dirty.workbook).toBe(true);
  });

  it("no-op when newName === name", async () => {
    const { bus } = await makeBus("02-multi-sheet.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:rename-sheet",
      payload: { name: "Expenses", newName: "Expenses" },
    });
    expect(m.status).toBe("approved");
  });

  it("rejects duplicate name (case-insensitive)", async () => {
    const { bus } = await makeBus("02-multi-sheet.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:rename-sheet",
      payload: { name: "Expenses", newName: "sales" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("duplicate-name");
  });

  it("rejects forbidden characters", async () => {
    const { bus } = await makeBus("02-multi-sheet.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:rename-sheet",
      payload: { name: "Expenses", newName: "Bad/Name" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-name");
  });

  it("rejects unknown source sheet", async () => {
    const { bus } = await makeBus("02-multi-sheet.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:rename-sheet",
      payload: { name: "Nope", newName: "Whatever" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-sheet");
  });
});

describe("CommandBus integration", () => {
  it("agent-sourced commands enter pending and can be approved", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "Z1", value: 7 },
      source: "agent",
      agentId: "test-agent",
    });
    expect(m.status).toBe("pending");
    expect(bus.getApproved().revision).toBe(0);
    expect(bus.getWorking().revision).toBe(1);
    bus.approveMutation(m.id);
    expect(bus.getApproved().revision).toBe(1);
  });

  it("CommandError thrown by handler surfaces as rejection (not throw)", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "🚫", value: 1 },
    });
    expect(m.status).toBe("rejected");
  });

  it("CommandError class is exported from core", () => {
    expect(CommandError).toBeDefined();
  });
});

function refOf(row: number, col: number): string {
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return `${s}${row + 1}`;
}
