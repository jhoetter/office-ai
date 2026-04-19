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
  it("adds a text box with TSV-rendered text on the active slide", async () => {
    const agent = await PptxAgent.empty();
    const before = agent.getSnapshot().root.slides[0]?.shapes.length ?? 0;
    await applyXlsxRangeToPptx({ agent, snapshot: SAMPLE, slideIndex: 0 });
    const slide = agent.getSnapshot().root.slides[0];
    expect(slide.shapes.length).toBe(before + 1);
    const inserted = slide.shapes[slide.shapes.length - 1];
    expect(inserted.kind).toBe("text");
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
