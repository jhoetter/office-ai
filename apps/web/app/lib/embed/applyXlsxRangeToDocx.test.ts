import { DocxAgent } from "@officeai/docx";
import type { XlsxClipboardSnapshot } from "@officeai/xlsx";
import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { applyXlsxRangeToDocx } from "./applyXlsxRangeToDocx";

function snapshot(rows: ReadonlyArray<ReadonlyArray<string | number | null>>): XlsxClipboardSnapshot {
  const height = rows.length;
  const width = Math.max(0, ...rows.map((r) => r.length));
  return {
    origin: { sheet: "Sheet1", range: `A1:${String.fromCharCode(64 + width)}${height}` },
    width,
    height,
    cells: rows.map((row) =>
      Array.from({ length: width }, (_, c) => {
        const v = row[c];
        if (v === null || v === undefined) return null;
        return { value: v };
      })
    ),
    merges: [],
  };
}

describe("applyXlsxRangeToDocx", () => {
  it("inserts a typed table and fills cells from the snapshot", async () => {
    const agent = await DocxAgent.empty({ idMinter: deterministicIdMinter() });
    const snap = snapshot([
      ["Name", "Score"],
      ["Ada", 99],
      ["Linus", 42],
    ]);
    await applyXlsxRangeToDocx({ agent, snapshot: snap, paragraphIndex: 0 });

    const doc = agent.getSnapshot().root;
    const table = doc.body.find((b) => b.kind === "table");
    expect(table).toBeDefined();
    if (table?.kind !== "table") throw new Error("expected a table");
    expect(table.rows).toHaveLength(3);
    expect(table.rows[0].cells).toHaveLength(2);
    const cellText = (r: number, c: number): string => {
      const cell = table.rows[r].cells[c];
      const para = cell.body.find((b) => b.kind === "paragraph");
      if (para?.kind !== "paragraph") return "";
      const run = para.children.find((ch) => ch.kind === "run");
      if (run?.kind !== "run") return "";
      const text = run.children.find((ch) => ch.kind === "text");
      return text?.kind === "text" ? text.text : "";
    };
    expect(cellText(0, 0)).toBe("Name");
    expect(cellText(0, 1)).toBe("Score");
    expect(cellText(1, 0)).toBe("Ada");
    expect(cellText(1, 1)).toBe("99");
    expect(cellText(2, 1)).toBe("42");
  });

  it("is a no-op for an empty snapshot", async () => {
    const agent = await DocxAgent.empty({ idMinter: deterministicIdMinter() });
    const before = agent.getSnapshot().root.body.length;
    await applyXlsxRangeToDocx({
      agent,
      snapshot: { origin: { sheet: "S", range: "A1:A1" }, width: 0, height: 0, cells: [], merges: [] },
      paragraphIndex: 0,
    });
    expect(agent.getSnapshot().root.body.length).toBe(before);
  });

  it("renders formula cells with a leading '=' as text", async () => {
    const agent = await DocxAgent.empty({ idMinter: deterministicIdMinter() });
    const snap: XlsxClipboardSnapshot = {
      origin: { sheet: "Sheet1", range: "A1:A1" },
      width: 1,
      height: 1,
      cells: [[{ value: 100, formula: "SUM(A1:A10)" }]],
      merges: [],
    };
    await applyXlsxRangeToDocx({ agent, snapshot: snap, paragraphIndex: 0 });
    const doc = agent.getSnapshot().root;
    const table = doc.body.find((b) => b.kind === "table");
    if (table?.kind !== "table") throw new Error("expected a table");
    const para = table.rows[0].cells[0].body[0];
    if (para?.kind !== "paragraph") throw new Error("expected a paragraph");
    const run = para.children[0];
    if (run?.kind !== "run") throw new Error("expected a run");
    const txt = run.children[0];
    expect(txt.kind === "text" && txt.text).toBe("=SUM(A1:A10)");
  });
});
