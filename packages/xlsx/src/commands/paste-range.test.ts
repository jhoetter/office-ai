import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { CommandBus, defaultIdMinter } from "@officeai/core";
import { extractClipboardSnapshot, type XlsxClipboardSnapshot } from "../clipboard/snapshot.js";
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

function literal(value: number | string | boolean | null): XlsxClipboardSnapshot {
  return {
    origin: { sheet: "ad-hoc", range: "A1" },
    width: 1,
    height: 1,
    cells: [[value === null ? null : { value }]],
    merges: [],
  };
}

describe("xlsx:paste-range — basic value paste", () => {
  it("writes a 2×2 block of values atomically", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const source: XlsxClipboardSnapshot = {
      origin: { sheet: "x", range: "A1:B2" },
      width: 2,
      height: 2,
      cells: [
        [{ value: 1 }, { value: 2 }],
        [{ value: 3 }, { value: 4 }],
      ],
      merges: [],
    };
    const m = await bus.dispatch({
      type: "xlsx:paste-range",
      payload: { sheet: "Inventory", target: "Z90", source },
    });
    expect(m.status).toBe("approved");
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(89, 25))?.value).toBe(1);
    expect(sheet.cells.get(cellKey(89, 26))?.value).toBe(2);
    expect(sheet.cells.get(cellKey(90, 25))?.value).toBe(3);
    expect(sheet.cells.get(cellKey(90, 26))?.value).toBe(4);
  });

  it("clears destination cells when source is null", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:paste-range",
      payload: { sheet: "Inventory", target: "Y1", source: literal("seed") },
    });
    expect(bus.getWorking().root.sheets[0].cells.get(cellKey(0, 24))?.value).toBe("seed");
    await bus.dispatch({
      type: "xlsx:paste-range",
      payload: { sheet: "Inventory", target: "Y1", source: literal(null) },
    });
    expect(bus.getWorking().root.sheets[0].cells.has(cellKey(0, 24))).toBe(false);
  });

  it("rejects an empty snapshot", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:paste-range",
      payload: {
        sheet: "Inventory",
        target: "A1",
        source: { origin: { sheet: "x", range: "A1" }, width: 0, height: 0, cells: [], merges: [] },
      },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-clipboard");
  });

  it("rejects a paste that runs off the right edge", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    // XFD is the last column (col 16383, 0-based). A single cell at
    // XFD1 fits; a 2-wide paste runs over the edge.
    const m = await bus.dispatch({
      type: "xlsx:paste-range",
      payload: { sheet: "Inventory", target: "XFD1", source: literal(1) },
    });
    expect(m.status).toBe("approved");
    const m2 = await bus.dispatch({
      type: "xlsx:paste-range",
      payload: {
        sheet: "Inventory",
        target: "XFD1",
        source: {
          origin: { sheet: "x", range: "A1:B1" },
          width: 2,
          height: 1,
          cells: [[{ value: 1 }, { value: 2 }]],
          merges: [],
        },
      },
    });
    expect(m2.status).toBe("rejected");
    expect(m2.rejection?.code).toBe("out-of-bounds");
  });
});

describe("xlsx:paste-range — formula relative shift", () => {
  it("shifts a relative ref by the destination delta", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    // Seed B1 = 10 so the source formula's reference resolves.
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "B1", value: 10 },
    });
    const source: XlsxClipboardSnapshot = {
      origin: { sheet: "Inventory", range: "A1" },
      width: 1,
      height: 1,
      // =B1 — pasted at A2 should become =B2 (relative shift by 1 row).
      cells: [[{ value: null, formula: "B1" }]],
      merges: [],
    };
    const m = await bus.dispatch({
      type: "xlsx:paste-range",
      payload: { sheet: "Inventory", target: "A2", source },
    });
    expect(m.status).toBe("approved");
    const sheet = bus.getWorking().root.sheets[0];
    const pasted = sheet.cells.get(cellKey(1, 0));
    expect(pasted?.formula?.text).toBe("B2");
  });

  it("preserves absolute refs ($A$1)", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "B1", value: 10 },
    });
    const source: XlsxClipboardSnapshot = {
      origin: { sheet: "Inventory", range: "A1" },
      width: 1,
      height: 1,
      cells: [[{ value: null, formula: "$B$1" }]],
      merges: [],
    };
    const m = await bus.dispatch({
      type: "xlsx:paste-range",
      payload: { sheet: "Inventory", target: "A5", source },
    });
    expect(m.status).toBe("approved");
    expect(bus.getWorking().root.sheets[0].cells.get(cellKey(4, 0))?.formula?.text).toBe("$B$1");
  });

  it("partially shifts mixed-absoluteness ($B1)", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "B5", value: 10 },
    });
    const source: XlsxClipboardSnapshot = {
      origin: { sheet: "Inventory", range: "A1" },
      width: 1,
      height: 1,
      // $B1 — col absolute, row relative; shift A1→A5 should give $B5.
      cells: [[{ value: null, formula: "$B1" }]],
      merges: [],
    };
    await bus.dispatch({
      type: "xlsx:paste-range",
      payload: { sheet: "Inventory", target: "A5", source },
    });
    const text = bus.getWorking().root.sheets[0].cells.get(cellKey(4, 0))?.formula?.text;
    expect(text).toBe("$B5");
  });

  it("recalculates pasted formulas immediately", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    // 2-row source: row 0 = literal 7, row 1 = formula `=A1*2` (which
    // in source-local coords means "value above me times 2"). After a
    // relative paste at C5 the formula's A1 ref shifts to C5 (the
    // upper neighbour of C6) and the recalc gives 14.
    const source: XlsxClipboardSnapshot = {
      origin: { sheet: "Inventory", range: "A1:A2" },
      width: 1,
      height: 2,
      cells: [[{ value: 7 }], [{ value: null, formula: "A1*2" }]],
      merges: [],
    };
    await bus.dispatch({
      type: "xlsx:paste-range",
      payload: { sheet: "Inventory", target: "C5", source },
    });
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(4, 2))?.value).toBe(7);
    expect(sheet.cells.get(cellKey(5, 2))?.formula?.text).toBe("C5*2");
    expect(sheet.cells.get(cellKey(5, 2))?.value).toBe(14);
  });
});

describe("xlsx:paste-range — modes", () => {
  it("values-only preserves destination styleId", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    // Apply a style at C3 first.
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "C3", value: "kept" },
    });
    await bus.dispatch({
      type: "xlsx:set-cell-format",
      payload: { sheet: "Inventory", range: "C3", format: { font: { bold: true } } },
    });
    const before = bus.getWorking().root.sheets[0].cells.get(cellKey(2, 2));
    const beforeStyleId = before?.styleId;
    expect(beforeStyleId).toBeTypeOf("number");

    const source: XlsxClipboardSnapshot = {
      origin: { sheet: "x", range: "A1" },
      width: 1,
      height: 1,
      cells: [[{ value: 999, styleId: 0 }]],
      merges: [],
    };
    await bus.dispatch({
      type: "xlsx:paste-range",
      payload: { sheet: "Inventory", target: "C3", source, mode: "values" },
    });
    const after = bus.getWorking().root.sheets[0].cells.get(cellKey(2, 2));
    expect(after?.value).toBe(999);
    expect(after?.styleId).toBe(beforeStyleId);
  });

  it("formats-only preserves destination value", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "D4", value: "keep-me" },
    });
    // Apply a style elsewhere then snapshot the styleId.
    await bus.dispatch({
      type: "xlsx:set-cell-format",
      payload: { sheet: "Inventory", range: "D5", format: { font: { italic: true } } },
    });
    const styleId = bus.getWorking().root.sheets[0].cells.get(cellKey(4, 3))?.styleId;
    expect(styleId).toBeTypeOf("number");

    const source: XlsxClipboardSnapshot = {
      origin: { sheet: "x", range: "A1" },
      width: 1,
      height: 1,
      cells: [[{ value: "ignored", styleId }]],
      merges: [],
    };
    await bus.dispatch({
      type: "xlsx:paste-range",
      payload: { sheet: "Inventory", target: "D4", source, mode: "formats" },
    });
    const after = bus.getWorking().root.sheets[0].cells.get(cellKey(3, 3));
    expect(after?.value).toBe("keep-me");
    expect(after?.styleId).toBe(styleId);
  });
});

describe("xlsx:paste-range — transpose + merges", () => {
  it("transposes a 1×3 row into a 3×1 column", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const source: XlsxClipboardSnapshot = {
      origin: { sheet: "x", range: "A1:C1" },
      width: 3,
      height: 1,
      cells: [[{ value: 1 }, { value: 2 }, { value: 3 }]],
      merges: [],
    };
    await bus.dispatch({
      type: "xlsx:paste-range",
      payload: { sheet: "Inventory", target: "F10", source, transpose: true },
    });
    const sheet = bus.getWorking().root.sheets[0];
    expect(sheet.cells.get(cellKey(9, 5))?.value).toBe(1);
    expect(sheet.cells.get(cellKey(10, 5))?.value).toBe(2);
    expect(sheet.cells.get(cellKey(11, 5))?.value).toBe(3);
  });

  it("round-trips a snapshot with merges", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:set-cell-value",
      payload: { sheet: "Inventory", ref: "P1", value: "hdr" },
    });
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "P1:Q1" },
    });
    const sheet = bus.getWorking().root.sheets[0];
    const snap = extractClipboardSnapshot(sheet, "P1:Q1");
    expect(snap.merges.length).toBe(1);

    await bus.dispatch({
      type: "xlsx:paste-range",
      payload: { sheet: "Inventory", target: "P20", source: snap },
    });
    const next = bus.getWorking().root.sheets[0];
    expect(next.cells.get(cellKey(19, 15))?.value).toBe("hdr");
    expect(next.merges.some((m) => m.r1 === 19 && m.c1 === 15 && m.r2 === 19 && m.c2 === 16)).toBe(true);
  });

  it("rejects partial overlap with an existing merge", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "S1:T2" },
    });
    // Target S1 is the merge anchor (passes assertNotMergedNonAnchor).
    // The pasted merge S1:U1 partially overlaps the existing S1:T2.
    const m = await bus.dispatch({
      type: "xlsx:paste-range",
      payload: {
        sheet: "Inventory",
        target: "S1",
        source: {
          origin: { sheet: "x", range: "A1:C1" },
          width: 3,
          height: 1,
          cells: [[{ value: "a" }, { value: "b" }, { value: "c" }]],
          merges: [{ r0: 0, c0: 0, r1: 0, c1: 2 }],
        },
      },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("merge-overlap");
  });
});

describe("extractClipboardSnapshot", () => {
  it("captures values + formulas + styleIds + merges", async () => {
    const { bus } = await makeBus("03-formulas-basic.xlsx");
    const sheet = bus.getWorking().root.sheets[0];
    const snap = extractClipboardSnapshot(sheet, "A1:B3");
    expect(snap.height).toBe(3);
    expect(snap.width).toBe(2);
    // shape sanity — cells matrix is 3×2 with sparse nulls
    expect(snap.cells.length).toBe(3);
    expect(snap.cells[0]?.length).toBe(2);
  });

  it("clips merges to the requested range (drops half-clipped)", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    await bus.dispatch({
      type: "xlsx:merge-cells",
      payload: { sheet: "Inventory", range: "X1:Y2" },
    });
    const sheet = bus.getWorking().root.sheets[0];
    // Capture only X1:X2 — half the merge → should drop it.
    const half = extractClipboardSnapshot(sheet, "X1:X2");
    expect(half.merges).toEqual([]);
    // Capture the full merge — should keep it.
    const full = extractClipboardSnapshot(sheet, "X1:Y2");
    expect(full.merges.length).toBe(1);
    expect(full.merges[0]).toEqual({ r0: 0, c0: 0, r1: 1, c1: 1 });
  });
});
