import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { PptxAgent } from "../agent/agent.js";
import type { TableShape } from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

function cellText(table: TableShape, row: number, col: number): string {
  const cell = table.rows[row]!.cells[col]!;
  return cell.txBody.paragraphs
    .map((p) => p.runs.map((r) => r.text).join(""))
    .join("\n");
}

describe("pptx:insert-table", () => {
  it("authors a real TableShape with row/cell structure on the slide", async () => {
    const buf = await readFile(join(FIXTURES_DIR.pathname, "04-multi-shape.pptx"));
    const agent = await PptxAgent.fromBuffer(buf);
    const before = agent.getSnapshot();
    const slide = before.root.slides[0]!;
    const shapesBefore = slide.shapes.length;

    await agent.applyCommand({
      type: "pptx:insert-table",
      payload: {
        slideIndex: 0,
        x: 1_000_000,
        y: 1_000_000,
        cx: 4_000_000,
        cy: 1_000_000,
        data: [
          ["Name", "Score"],
          ["Ada", 99],
          ["Linus", 87.5],
        ],
      },
    });

    const after = agent.getSnapshot();
    const slideAfter = after.root.slides[0]!;
    expect(slideAfter.shapes.length).toBe(shapesBefore + 1);
    const table = slideAfter.shapes[slideAfter.shapes.length - 1] as TableShape;
    expect(table.kind).toBe("table");
    expect(table.rows).toHaveLength(3);
    expect(table.columnWidths).toHaveLength(2);
    // Even split + remainder onto the last column should sum to cx.
    expect(table.columnWidths.reduce((a, b) => a + b, 0)).toBe(4_000_000);
    expect(table.rows.reduce((a, r) => a + r.height, 0)).toBe(1_000_000);
    expect(cellText(table, 0, 0)).toBe("Name");
    expect(cellText(table, 1, 0)).toBe("Ada");
    expect(cellText(table, 2, 1)).toBe("87.5");
    expect(after.dirty.slides.has(slide.partPath)).toBe(true);
  });

  it("survives a serialize → reparse roundtrip preserving cell text", async () => {
    const buf = await readFile(join(FIXTURES_DIR.pathname, "04-multi-shape.pptx"));
    const agent = await PptxAgent.fromBuffer(buf);
    await agent.applyCommand({
      type: "pptx:insert-table",
      payload: {
        slideIndex: 0,
        x: 500_000,
        y: 500_000,
        data: [
          ["A", "B", "C"],
          [1, 2, 3],
        ],
      },
    });
    const out = await agent.exportFile();
    const reagent = await PptxAgent.fromBuffer(new Uint8Array(out));
    const slide = reagent.getSnapshot().root.slides[0]!;
    const table = slide.shapes.find((s): s is TableShape => s.kind === "table");
    expect(table).toBeDefined();
    expect(table!.rows).toHaveLength(2);
    expect(table!.columnWidths).toHaveLength(3);
    expect(cellText(table!, 0, 1)).toBe("B");
    expect(cellText(table!, 1, 2)).toBe("3");
  });

  it("rejects empty grids", async () => {
    const buf = await readFile(join(FIXTURES_DIR.pathname, "04-multi-shape.pptx"));
    const agent = await PptxAgent.fromBuffer(buf);
    const result = await agent.applyCommand({
      type: "pptx:insert-table",
      payload: {
        slideIndex: 0,
        x: 0,
        y: 0,
        data: [],
      },
    });
    expect(result.status).toBe("rejected");
    if (result.status === "rejected") {
      expect(result.rejection.message).toMatch(/at least one row/);
    }
  });
});
