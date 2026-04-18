import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { XlsxAgent } from "./agent.js";
import { cellKey, parseA1 } from "../model/refs.js";

/**
 * Per-handler undo round-trip coverage.
 *
 * Each test (a) captures the relevant slice of the snapshot,
 * (b) dispatches a P0 / P13 command, (c) calls `undo()` and
 * asserts the slice round-trips back to its original state.
 * This is the fastest way to flush out handlers whose `before`
 * snapshot accidentally shares mutable state with `after`.
 */
describe("XlsxAgent undo round-trips", () => {
  async function newAgent() {
    return XlsxAgent.empty({ idMinter: deterministicIdMinter("undo-") });
  }

  function valueAt(a: XlsxAgent, sheetName: string, ref: string): unknown {
    const snap = a.getSnapshot();
    const sheet = snap.root.sheets.find((s) => s.name === sheetName)!;
    const { row, col } = parseA1(ref);
    return sheet.cells.get(cellKey(row, col))?.value;
  }

  function formulaAt(a: XlsxAgent, sheetName: string, ref: string): string | undefined {
    const snap = a.getSnapshot();
    const sheet = snap.root.sheets.find((s) => s.name === sheetName)!;
    const { row, col } = parseA1(ref);
    return sheet.cells.get(cellKey(row, col))?.formula;
  }

  it("undoes set-cell-value", async () => {
    const a = await newAgent();
    const sheet = a.listSheets()[0]!.name;
    await a.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet, ref: "B2", value: 7 },
    });
    expect(valueAt(a, sheet, "B2")).toBe(7);
    a.undo();
    expect(valueAt(a, sheet, "B2")).toBeUndefined();
    expect(a.canRedo()).toBe(true);
    a.redo();
    expect(valueAt(a, sheet, "B2")).toBe(7);
  });

  it("undoes set-cell-formula", async () => {
    const a = await newAgent();
    const sheet = a.listSheets()[0]!.name;
    await a.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet, ref: "A1", value: 4 },
    });
    await a.applyCommand({
      type: "xlsx:set-cell-formula",
      payload: { sheet, ref: "B1", formula: "=A1*3" },
    });
    expect(valueAt(a, sheet, "B1")).toBe(12);
    a.undo();
    expect(formulaAt(a, sheet, "B1")).toBeUndefined();
  });

  it("undoes paste-range", async () => {
    const a = await newAgent();
    const sheet = a.listSheets()[0]!.name;
    await a.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet, ref: "A1", value: 1 },
    });
    const snap = a.getClipboardSnapshot({ sheet, range: "A1:A1" });
    await a.applyCommand({
      type: "xlsx:paste-range",
      payload: { sheet, target: "C3", source: snap, mode: "all" },
    });
    expect(valueAt(a, sheet, "C3")).toBe(1);
    a.undo();
    expect(valueAt(a, sheet, "C3")).toBeUndefined();
  });

  it("undoes fill-range", async () => {
    const a = await newAgent();
    const sheet = a.listSheets()[0]!.name;
    await a.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet, ref: "A1", value: 1 },
    });
    await a.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet, ref: "A2", value: 2 },
    });
    await a.applyCommand({
      type: "xlsx:fill-range",
      payload: {
        sheet,
        source: "A1:A2",
        target: "A1:A5",
        direction: "down",
      },
    });
    expect(valueAt(a, sheet, "A5")).toBe(5);
    a.undo();
    expect(valueAt(a, sheet, "A5")).toBeUndefined();
    expect(valueAt(a, sheet, "A2")).toBe(2);
  });

  it("undoes text-to-columns", async () => {
    const a = await newAgent();
    const sheet = a.listSheets()[0]!.name;
    await a.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet, ref: "A1", value: "a,b,c" },
    });
    await a.applyCommand({
      type: "xlsx:text-to-columns",
      payload: { sheet, range: "A1:A1", delimiter: "," },
    });
    expect(valueAt(a, sheet, "C1")).toBe("c");
    a.undo();
    expect(valueAt(a, sheet, "A1")).toBe("a,b,c");
    expect(valueAt(a, sheet, "B1")).toBeUndefined();
  });

  it("undoes merge-cells", async () => {
    const a = await newAgent();
    const sheet = a.listSheets()[0]!.name;
    await a.applyCommand({
      type: "xlsx:merge-cells",
      payload: { sheet, range: "A1:B2" },
    });
    expect(a.getSnapshot().root.sheets[0]!.merges.length).toBe(1);
    a.undo();
    expect(a.getSnapshot().root.sheets[0]!.merges.length).toBe(0);
  });

  it("undoes insert-row", async () => {
    const a = await newAgent();
    const sheet = a.listSheets()[0]!.name;
    await a.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet, ref: "A1", value: "top" },
    });
    await a.applyCommand({
      type: "xlsx:insert-row",
      payload: { sheet, at: 1, count: 1 },
    });
    expect(valueAt(a, sheet, "A2")).toBe("top");
    a.undo();
    expect(valueAt(a, sheet, "A1")).toBe("top");
    expect(valueAt(a, sheet, "A2")).toBeUndefined();
  });

  it("clears the redo stack after a fresh authored mutation", async () => {
    const a = await newAgent();
    const sheet = a.listSheets()[0]!.name;
    await a.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet, ref: "A1", value: 1 },
    });
    a.undo();
    expect(a.canRedo()).toBe(true);
    await a.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet, ref: "A1", value: 99 },
    });
    expect(a.canRedo()).toBe(false);
    expect(valueAt(a, sheet, "A1")).toBe(99);
  });
});
