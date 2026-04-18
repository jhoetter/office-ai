import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { PptxAgent } from "../agent/agent.js";
import type { TableShape } from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadTableAgent(): Promise<{
  agent: PptxAgent;
  table: TableShape;
}> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, "06-with-table.pptx"));
  const agent = await PptxAgent.fromBuffer(buf);
  const slide = agent.getSnapshot().root.slides[0]!;
  const table = slide.shapes.find((s): s is TableShape => s.kind === "table")!;
  expect(table).toBeDefined();
  return { agent, table };
}

function flattenCellText(cell: TableShape["rows"][number]["cells"][number]): string {
  return cell.txBody.paragraphs
    .flatMap((p) => p.runs.filter((r) => !r.isLineBreak).map((r) => r.text))
    .join("");
}

function getTable(agent: PptxAgent): TableShape {
  const slide = agent.getSnapshot().root.slides[0]!;
  return slide.shapes.find((s): s is TableShape => s.kind === "table")!;
}

describe("F2: pptx:table-set-cell-text", () => {
  it("replaces a cell's text and marks the slide dirty", async () => {
    const { agent, table } = await loadTableAgent();
    await agent.applyCommand({
      type: "pptx:table-set-cell-text",
      payload: { slideIndex: 0, shapeId: table.id, row: 1, column: 1, text: "Edited 🎉" },
    });
    const updated = getTable(agent);
    const cell = updated.rows[1]!.cells[1]!;
    expect(flattenCellText(cell)).toBe("Edited 🎉");
    expect(agent.getSnapshot().dirty.slides.size).toBe(1);
  });

  it("rejects out-of-range row / column", async () => {
    const { agent, table } = await loadTableAgent();
    const m1 = await agent.applyCommand({
      type: "pptx:table-set-cell-text",
      payload: { slideIndex: 0, shapeId: table.id, row: 99, column: 0, text: "" },
    });
    expect(m1.rejection?.code).toBe("unknown-target");
    const m2 = await agent.applyCommand({
      type: "pptx:table-set-cell-text",
      payload: { slideIndex: 0, shapeId: table.id, row: 0, column: 99, text: "" },
    });
    expect(m2.rejection?.code).toBe("unknown-target");
  });

  it("rejects non-table shapes with not-applicable", async () => {
    const { agent } = await loadTableAgent();
    const slide = agent.getSnapshot().root.slides[0]!;
    const text = slide.shapes.find((s) => s.kind === "text")!;
    const m = await agent.applyCommand({
      type: "pptx:table-set-cell-text",
      payload: { slideIndex: 0, shapeId: text.id, row: 0, column: 0, text: "x" },
    });
    expect(m.rejection?.code).toBe("not-applicable");
  });
});

describe("F2: pptx:table-add-row", () => {
  it("appends a row with empty cells when at is omitted", async () => {
    const { agent, table } = await loadTableAgent();
    const before = table.rows.length;
    await agent.applyCommand({
      type: "pptx:table-add-row",
      payload: { slideIndex: 0, shapeId: table.id },
    });
    const updated = getTable(agent);
    expect(updated.rows.length).toBe(before + 1);
    const last = updated.rows[updated.rows.length - 1]!;
    expect(last.cells.length).toBe(updated.columnWidths.length);
    for (const cell of last.cells) expect(flattenCellText(cell)).toBe("");
  });

  it("inserts at the requested position and uses default median height", async () => {
    const { agent, table } = await loadTableAgent();
    await agent.applyCommand({
      type: "pptx:table-add-row",
      payload: { slideIndex: 0, shapeId: table.id, at: 1, height: 500000 },
    });
    const updated = getTable(agent);
    expect(updated.rows[1]!.height).toBe(500000);
  });

  it("rejects invalid-position when at is out of range", async () => {
    const { agent, table } = await loadTableAgent();
    const m = await agent.applyCommand({
      type: "pptx:table-add-row",
      payload: { slideIndex: 0, shapeId: table.id, at: -1 },
    });
    expect(m.rejection?.code).toBe("invalid-position");
  });
});

describe("F2: pptx:table-delete-row", () => {
  it("removes the requested row", async () => {
    const { agent, table } = await loadTableAgent();
    const before = table.rows.length;
    await agent.applyCommand({
      type: "pptx:table-delete-row",
      payload: { slideIndex: 0, shapeId: table.id, row: 0 },
    });
    expect(getTable(agent).rows.length).toBe(before - 1);
  });

  it("rejects deleting the last row", async () => {
    const { agent, table } = await loadTableAgent();
    let current = table;
    while (current.rows.length > 1) {
      await agent.applyCommand({
        type: "pptx:table-delete-row",
        payload: { slideIndex: 0, shapeId: current.id, row: 0 },
      });
      current = getTable(agent);
    }
    const m = await agent.applyCommand({
      type: "pptx:table-delete-row",
      payload: { slideIndex: 0, shapeId: current.id, row: 0 },
    });
    expect(m.rejection?.code).toBe("invalid-payload");
  });
});

describe("F2: pptx:table-add-column", () => {
  it("appends a column to columnWidths and to every row", async () => {
    const { agent, table } = await loadTableAgent();
    const beforeCols = table.columnWidths.length;
    const beforeRowCells = table.rows[0]!.cells.length;
    await agent.applyCommand({
      type: "pptx:table-add-column",
      payload: { slideIndex: 0, shapeId: table.id, width: 999999 },
    });
    const updated = getTable(agent);
    expect(updated.columnWidths.length).toBe(beforeCols + 1);
    expect(updated.columnWidths[updated.columnWidths.length - 1]).toBe(999999);
    for (const row of updated.rows) {
      expect(row.cells.length).toBe(beforeRowCells + 1);
      const newCell = row.cells[row.cells.length - 1]!;
      expect(flattenCellText(newCell)).toBe("");
    }
  });

  it("rejects invalid-position", async () => {
    const { agent, table } = await loadTableAgent();
    const m = await agent.applyCommand({
      type: "pptx:table-add-column",
      payload: { slideIndex: 0, shapeId: table.id, at: 999 },
    });
    expect(m.rejection?.code).toBe("invalid-position");
  });
});

describe("F2: pptx:table-delete-column", () => {
  it("removes the requested column from columnWidths and every row", async () => {
    const { agent, table } = await loadTableAgent();
    const beforeCols = table.columnWidths.length;
    await agent.applyCommand({
      type: "pptx:table-delete-column",
      payload: { slideIndex: 0, shapeId: table.id, column: 0 },
    });
    const updated = getTable(agent);
    expect(updated.columnWidths.length).toBe(beforeCols - 1);
    for (const row of updated.rows) {
      expect(row.cells.length).toBe(beforeCols - 1);
    }
  });

  it("rejects deleting the last column", async () => {
    const { agent, table } = await loadTableAgent();
    let current = table;
    while (current.columnWidths.length > 1) {
      await agent.applyCommand({
        type: "pptx:table-delete-column",
        payload: { slideIndex: 0, shapeId: current.id, column: 0 },
      });
      current = getTable(agent);
    }
    const m = await agent.applyCommand({
      type: "pptx:table-delete-column",
      payload: { slideIndex: 0, shapeId: current.id, column: 0 },
    });
    expect(m.rejection?.code).toBe("invalid-payload");
  });
});

describe("F2: roundtrip after table edits", () => {
  it("survives parse → edit → serialize → parse with edits intact", async () => {
    const { agent, table } = await loadTableAgent();
    await agent.applyCommand({
      type: "pptx:table-set-cell-text",
      payload: { slideIndex: 0, shapeId: table.id, row: 0, column: 0, text: "Quarter (edited)" },
    });
    await agent.applyCommand({
      type: "pptx:table-add-row",
      payload: { slideIndex: 0, shapeId: table.id },
    });
    await agent.applyCommand({
      type: "pptx:table-add-column",
      payload: { slideIndex: 0, shapeId: table.id },
    });
    const out = await agent.exportFile();
    const reloaded = await PptxAgent.fromBuffer(out);
    const reTable = getTable(reloaded);
    expect(flattenCellText(reTable.rows[0]!.cells[0]!)).toBe("Quarter (edited)");
    expect(reTable.rows.length).toBe(table.rows.length + 1);
    expect(reTable.columnWidths.length).toBe(table.columnWidths.length + 1);
  });
});
