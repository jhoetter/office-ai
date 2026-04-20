import { PptxAgent } from "@officeai/pptx/agent";
import type { TableShape } from "@officeai/pptx";
import type { XlsxClipboardCell, XlsxClipboardSnapshot } from "@officeai/xlsx";
import { describe, expect, it, vi } from "vitest";
import { applyXlsxRangeToPptx } from "./applyXlsxRangeToPptx";

const SAMPLE: XlsxClipboardSnapshot = {
  origin: { sheet: "Sheet1", range: "A1:C2" },
  width: 3,
  height: 2,
  cells: [
    [{ value: "Name" }, { value: "Score" }, { value: "Note" }],
    [{ value: "Ada" }, { value: 99 }, { value: true }],
  ],
  merges: [],
};

function findFirstTable(agent: PptxAgent): TableShape | null {
  for (const slide of agent.getSnapshot().root.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind === "table") return shape;
    }
  }
  return null;
}

function cellText(table: TableShape, r: number, c: number): string {
  const cell = table.rows[r]?.cells[c];
  if (!cell) return "";
  return cell.txBody.paragraphs
    .flatMap((p) => p.runs.filter((run) => !run.isLineBreak).map((run) => run.text))
    .join("\n");
}

describe("applyXlsxRangeToPptx", () => {
  it("dispatches a single pptx:insert-table with rows/cols/cells from the snapshot", async () => {
    const agent = await PptxAgent.empty();
    const before = agent.getSnapshot().root.slides[0]?.shapes.length ?? 0;

    await applyXlsxRangeToPptx({ agent, snapshot: SAMPLE, slideIndex: 0 });

    const slide = agent.getSnapshot().root.slides[0];
    expect(slide.shapes.length).toBe(before + 1);
    const table = findFirstTable(agent);
    expect(table).not.toBeNull();
    if (!table) throw new Error("expected a table");
    expect(table.rows).toHaveLength(2);
    expect(table.columnWidths).toHaveLength(3);
    expect(cellText(table, 0, 0)).toBe("Name");
    expect(cellText(table, 0, 1)).toBe("Score");
    expect(cellText(table, 0, 2)).toBe("Note");
    expect(cellText(table, 1, 0)).toBe("Ada");
    expect(cellText(table, 1, 1)).toBe("99");
    expect(cellText(table, 1, 2)).toBe("TRUE");
  });

  it("renders a formula cell's cached result when one is present", async () => {
    const agent = await PptxAgent.empty();
    const snap: XlsxClipboardSnapshot = {
      origin: { sheet: "Sheet1", range: "A1:A1" },
      width: 1,
      height: 1,
      cells: [[{ value: 42, formula: "SUM(A1:A10)" }]],
      merges: [],
    };
    await applyXlsxRangeToPptx({ agent, snapshot: snap, slideIndex: 0 });
    const table = findFirstTable(agent);
    if (!table) throw new Error("expected a table");
    expect(cellText(table, 0, 0)).toBe("42");
  });

  it("falls back to the raw formula text when no cached result is present", async () => {
    const agent = await PptxAgent.empty();
    const snap: XlsxClipboardSnapshot = {
      origin: { sheet: "Sheet1", range: "A1:A1" },
      width: 1,
      height: 1,
      cells: [[{ value: null, formula: "SUM(A1:A10)" }]],
      merges: [],
    };
    await applyXlsxRangeToPptx({ agent, snapshot: snap, slideIndex: 0 });
    const table = findFirstTable(agent);
    if (!table) throw new Error("expected a table");
    expect(cellText(table, 0, 0)).toBe("=SUM(A1:A10)");
  });

  it("emits empty strings (not 'undefined') for null cells", async () => {
    const agent = await PptxAgent.empty();
    const snap: XlsxClipboardSnapshot = {
      origin: { sheet: "Sheet1", range: "A1:B2" },
      width: 2,
      height: 2,
      cells: [
        [{ value: "x" }, null],
        [null, { value: "y" }],
      ],
      merges: [],
    };
    await applyXlsxRangeToPptx({ agent, snapshot: snap, slideIndex: 0 });
    const table = findFirstTable(agent);
    if (!table) throw new Error("expected a table");
    expect(cellText(table, 0, 0)).toBe("x");
    expect(cellText(table, 0, 1)).toBe("");
    expect(cellText(table, 1, 0)).toBe("");
    expect(cellText(table, 1, 1)).toBe("y");
    // The table command renders an empty paragraph for empty cells; verify
    // that paragraph carries no run text (NOT the literal "undefined").
    expect(cellText(table, 0, 1)).not.toContain("undefined");
    expect(cellText(table, 1, 0)).not.toContain("undefined");
  });

  it("is a no-op for an empty snapshot", async () => {
    const agent = await PptxAgent.empty();
    const before = agent.getSnapshot().root.slides[0]?.shapes.length ?? 0;
    await applyXlsxRangeToPptx({
      agent,
      snapshot: { origin: { sheet: "S", range: "A1:A1" }, width: 0, height: 0, cells: [], merges: [] },
      slideIndex: 0,
    });
    expect(agent.getSnapshot().root.slides[0]?.shapes.length ?? 0).toBe(before);
  });

  it("propagates a rejection from the agent as a thrown error", async () => {
    const agent = await PptxAgent.empty();
    const spy = vi.spyOn(agent, "applyCommand").mockResolvedValue({
      accepted: false,
      rejection: { code: "invalid-payload", message: "rows must be a positive integer" },
      revision: agent.getSnapshot().revision,
    } as unknown as Awaited<ReturnType<typeof agent.applyCommand>>);

    await expect(
      applyXlsxRangeToPptx({ agent, snapshot: SAMPLE, slideIndex: 0 }),
    ).rejects.toThrow(/pptx:insert-table rejected: invalid-payload/);

    spy.mockRestore();
  });
});
