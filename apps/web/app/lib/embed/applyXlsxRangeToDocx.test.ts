import { DocxAgent } from "@officeai/docx";
import type { EffectiveStyle, XlsxClipboardCell, XlsxClipboardSnapshot } from "@officeai/xlsx";
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

  describe("style propagation", () => {
    /**
     * Pull the first run on the (r,c) cell of the first table in the
     * agent's snapshot. Returns `null` when the cell is empty (the
     * paste handler skips cells without text), so callers can assert
     * "no run / no run-props injected".
     */
    function firstRunPropsAt(
      agent: DocxAgent,
      r: number,
      c: number,
    ): { props: import("@officeai/docx").RunProperties; text: string } | null {
      const doc = agent.getSnapshot().root;
      const table = doc.body.find((b) => b.kind === "table");
      if (table?.kind !== "table") return null;
      const cell = table.rows[r]?.cells[c];
      if (!cell) return null;
      const para = cell.body.find((b) => b.kind === "paragraph");
      if (para?.kind !== "paragraph") return null;
      const run = para.children.find((ch) => ch.kind === "run");
      if (run?.kind !== "run") return null;
      const text = run.children.find((ch) => ch.kind === "text");
      if (!text || text.kind !== "text") return null;
      return { props: run.properties, text: text.text };
    }

    function styledCell(
      value: string | number,
      font: Omit<Partial<EffectiveStyle["font"]>, "opaqueExtras">,
    ): XlsxClipboardCell {
      const eff: EffectiveStyle = {
        numFmtId: 0,
        font: { ...font, opaqueExtras: [] },
        fill: { kind: "pattern", patternType: "none" },
        border: { left: {}, right: {}, top: {}, bottom: {}, diagonal: {}, opaqueExtras: [] },
      };
      return { value, effectiveStyle: eff };
    }

    function styledSnapshot(
      rows: ReadonlyArray<ReadonlyArray<XlsxClipboardCell | null>>,
    ): XlsxClipboardSnapshot {
      const height = rows.length;
      const width = Math.max(0, ...rows.map((r) => r.length));
      return {
        origin: { sheet: "S", range: `A1:${String.fromCharCode(64 + width)}${height}` },
        width,
        height,
        cells: rows.map((row) => Array.from({ length: width }, (_, c) => row[c] ?? null)),
        merges: [],
      };
    }

    it("propagates header style and body style into run properties", async () => {
      const agent = await DocxAgent.empty({ idMinter: deterministicIdMinter() });
      const snap = styledSnapshot([
        [
          // header cell: bold red 14pt Calibri
          styledCell("Name", {
            name: "Calibri",
            size: 14,
            bold: true,
            color: { rgb: "FFFF0000" },
          }),
          styledCell("Score", {
            name: "Calibri",
            size: 14,
            bold: true,
            color: { rgb: "FFFF0000" },
          }),
        ],
        [
          // body cell: italic blue
          styledCell("Ada", { italic: true, color: { rgb: "FF0000FF" } }),
          styledCell(99, { italic: true, color: { rgb: "FF0000FF" } }),
        ],
      ]);
      await applyXlsxRangeToDocx({ agent, snapshot: snap, paragraphIndex: 0 });

      const head = firstRunPropsAt(agent, 0, 0);
      expect(head?.text).toBe("Name");
      expect(head?.props.bold).toBe(true);
      expect(head?.props.fontFamily).toBe("Calibri");
      expect(head?.props.fontSize).toBe(28);
      expect(head?.props.color).toBe("FF0000");

      const body = firstRunPropsAt(agent, 1, 0);
      expect(body?.text).toBe("Ada");
      expect(body?.props.italic).toBe(true);
      expect(body?.props.color).toBe("0000FF");
      expect(body?.props.bold).toBeUndefined();
      expect(body?.props.fontSize).toBeUndefined();
      expect(body?.props.fontFamily).toBeUndefined();
    });

    it("auto-bolds the header row when the source style omits bold", async () => {
      const agent = await DocxAgent.empty({ idMinter: deterministicIdMinter() });
      const snap = styledSnapshot([
        [styledCell("Header", { name: "Arial" })],
        [styledCell("Body", { name: "Arial" })],
      ]);
      await applyXlsxRangeToDocx({ agent, snapshot: snap, paragraphIndex: 0 });

      const head = firstRunPropsAt(agent, 0, 0);
      expect(head?.props.bold).toBe(true);
      expect(head?.props.fontFamily).toBe("Arial");

      const body = firstRunPropsAt(agent, 1, 0);
      expect(body?.props.bold).toBeUndefined();
      expect(body?.props.fontFamily).toBe("Arial");
    });

    it("honours an explicit bold:false on a header cell (no auto-bold override)", async () => {
      const agent = await DocxAgent.empty({ idMinter: deterministicIdMinter() });
      const snap = styledSnapshot([[styledCell("Header", { bold: false })]]);
      await applyXlsxRangeToDocx({ agent, snapshot: snap, paragraphIndex: 0 });

      const head = firstRunPropsAt(agent, 0, 0);
      expect(head?.props.bold).toBe(false);
    });

    it("emits empty run properties for cells without an effective style (no defaults injected)", async () => {
      const agent = await DocxAgent.empty({ idMinter: deterministicIdMinter() });
      // Two-row snapshot so row 1 is unambiguously a body row that
      // should NOT be auto-bolded.
      const snap = snapshot([
        ["Header", "Other"],
        ["Body", "Cell"],
      ]);
      await applyXlsxRangeToDocx({ agent, snapshot: snap, paragraphIndex: 0 });

      // Header row (row 0) still auto-bolds even without a source
      // style — that's the documented header convention.
      const head = firstRunPropsAt(agent, 0, 0);
      expect(head?.props.bold).toBe(true);
      expect(head?.props.fontFamily).toBeUndefined();
      expect(head?.props.fontSize).toBeUndefined();
      expect(head?.props.color).toBeUndefined();
      expect(head?.props.italic).toBeUndefined();

      // Body row (row 1) gets a totally empty RunProperties bag.
      const body = firstRunPropsAt(agent, 1, 0);
      expect(body?.props).toEqual({});
    });
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
