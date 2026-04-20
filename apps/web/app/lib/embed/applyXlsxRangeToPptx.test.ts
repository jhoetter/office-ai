import { PptxAgent } from "@officeai/pptx/agent";
import type { XlsxClipboardSnapshot } from "@officeai/xlsx";
import { describe, expect, it } from "vitest";
import { applyXlsxRangeToPptx } from "./applyXlsxRangeToPptx";

const SAMPLE: XlsxClipboardSnapshot = {
  origin: { sheet: "Sheet1", range: "A1:B2" },
  width: 2,
  height: 2,
  cells: [
    [{ value: "Name" }, { value: "Score" }],
    [{ value: "Ada" }, { value: 99 }],
  ],
  merges: [],
};

describe("applyXlsxRangeToPptx", () => {
  it("adds a real TableShape on the active slide", async () => {
    const agent = await PptxAgent.empty();
    const before = agent.getSnapshot().root.slides[0]?.shapes.length ?? 0;
    await applyXlsxRangeToPptx({ agent, snapshot: SAMPLE, slideIndex: 0 });
    const slide = agent.getSnapshot().root.slides[0];
    expect(slide.shapes.length).toBe(before + 1);
    const inserted = slide.shapes[slide.shapes.length - 1];
    expect(inserted.kind).toBe("table");
    if (inserted.kind === "table") {
      expect(inserted.rows).toHaveLength(2);
      expect(inserted.columnWidths).toHaveLength(2);
      expect(inserted.rows[0]?.cells).toHaveLength(2);
      const cellText = (cellIdx: { r: number; c: number }) => {
        const cell = inserted.rows[cellIdx.r]!.cells[cellIdx.c]!;
        return cell.txBody.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n");
      };
      expect(cellText({ r: 0, c: 0 })).toBe("Name");
      expect(cellText({ r: 1, c: 1 })).toBe("99");
    }
  });

  it("inserts an OLE-embedded spreadsheet when mode='live'", async () => {
    const agent = await PptxAgent.empty();
    await applyXlsxRangeToPptx({ agent, snapshot: SAMPLE, slideIndex: 0, mode: "live" });
    const slide = agent.getSnapshot().root.slides[0];
    const inserted = slide.shapes[slide.shapes.length - 1];
    expect(inserted.kind).toBe("ole-spreadsheet");
  });

  it("inserts a chart when mode='chart'", async () => {
    const agent = await PptxAgent.empty();
    await applyXlsxRangeToPptx({
      agent,
      snapshot: {
        origin: { sheet: "Sheet1", range: "A1:C3" },
        width: 3,
        height: 3,
        cells: [
          [{ value: "" }, { value: "EU" }, { value: "US" }],
          [{ value: "Q1" }, { value: 10 }, { value: 20 }],
          [{ value: "Q2" }, { value: 15 }, { value: 25 }],
        ],
        merges: [],
      },
      slideIndex: 0,
      mode: "chart",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const inserted = slide.shapes[slide.shapes.length - 1];
    expect(inserted.kind).toBe("chart");
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
});
