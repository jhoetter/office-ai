import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { DEFAULT_DOC_ROOT_ATTRS, escapeXml, makeSyntheticDocx } from "../test-utils/synthetic.js";
import type { BlockNode, Table } from "../model/types.js";

function syntheticTableDoc(rows: ReadonlyArray<ReadonlyArray<string>>): string {
  const cols = rows[0]?.length ?? 0;
  const widths = new Array(cols).fill(2000);
  const grid = widths.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  const trXml = rows
    .map(
      (row) =>
        `<w:tr>${row
          .map(
            (text, c) =>
              `<w:tc><w:tcPr><w:tcW w:w="${widths[c]}" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:tc>`
          )
          .join("")}</w:tr>`
    )
    .join("");
  const tableXml = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${trXml}</w:tbl>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>${tableXml}<w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
}

async function loadTable(rows: ReadonlyArray<ReadonlyArray<string>>): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: syntheticTableDoc(rows) });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function tableBlock(snap: { root: { body: ReadonlyArray<BlockNode> } }): Table {
  for (const b of snap.root.body) {
    if (b.kind === "table") return b;
  }
  throw new Error("no table in body");
}

describe("docx:set-cell-shading", () => {
  it("stamps fill onto a single cell", async () => {
    const agent = await loadTable([
      ["A", "B"],
      ["C", "D"],
    ]);
    const id = tableBlock(agent.getSnapshot()).id;
    await agent.applyCommand({
      type: "docx:set-cell-shading",
      payload: { tableId: id, row: 0, column: 1, fill: "FFE699" },
    });
    const cell = tableBlock(agent.getSnapshot()).rows[0].cells[1];
    expect(cell.properties.shd?.fill).toBe("FFE699");
  });

  it("clears shading when fill is null", async () => {
    const agent = await loadTable([["A"], ["B"]]);
    const id = tableBlock(agent.getSnapshot()).id;
    await agent.applyCommand({
      type: "docx:set-cell-shading",
      payload: { tableId: id, row: 0, column: 0, fill: "112233" },
    });
    await agent.applyCommand({
      type: "docx:set-cell-shading",
      payload: { tableId: id, row: 0, column: 0, fill: null },
    });
    const cell = tableBlock(agent.getSnapshot()).rows[0].cells[0];
    expect(cell.properties.shd).toBeUndefined();
  });

  it("rejects malformed hex", async () => {
    const agent = await loadTable([["A"]]);
    const id = tableBlock(agent.getSnapshot()).id;
    const result = await agent.applyCommand({
      type: "docx:set-cell-shading",
      payload: { tableId: id, row: 0, column: 0, fill: "not-a-color" },
    });
    expect(result.status).toBe("rejected");
  });
});

describe("docx:set-cell-alignment", () => {
  it("stamps and clears vAlign", async () => {
    const agent = await loadTable([["A"]]);
    const id = tableBlock(agent.getSnapshot()).id;
    await agent.applyCommand({
      type: "docx:set-cell-alignment",
      payload: { tableId: id, row: 0, column: 0, vAlign: "center" },
    });
    expect(tableBlock(agent.getSnapshot()).rows[0].cells[0].properties.vAlign).toBe("center");

    await agent.applyCommand({
      type: "docx:set-cell-alignment",
      payload: { tableId: id, row: 0, column: 0, vAlign: null },
    });
    expect(tableBlock(agent.getSnapshot()).rows[0].cells[0].properties.vAlign).toBeUndefined();
  });
});

describe("docx:set-row-height", () => {
  it("stamps trHeight with rule", async () => {
    const agent = await loadTable([["A"], ["B"]]);
    const id = tableBlock(agent.getSnapshot()).id;
    await agent.applyCommand({
      type: "docx:set-row-height",
      payload: { tableId: id, row: 1, heightTwips: 720, rule: "exact" },
    });
    const props = tableBlock(agent.getSnapshot()).rows[1].properties;
    expect(props.trHeight?.value).toBe(720);
    expect(props.trHeight?.rule).toBe("exact");
  });
});

describe("docx:set-column-width", () => {
  it("updates the grid + matching tcW", async () => {
    const agent = await loadTable([
      ["A", "B"],
      ["C", "D"],
    ]);
    const id = tableBlock(agent.getSnapshot()).id;
    await agent.applyCommand({
      type: "docx:set-column-width",
      payload: { tableId: id, column: 0, widthTwips: 3500 },
    });
    const t = tableBlock(agent.getSnapshot());
    expect(t.grid[0].w).toBe(3500);
    for (const row of t.rows) {
      expect(row.cells[0].properties.tcW?.value).toBe(3500);
    }
  });
});

describe("docx:merge-cells-horizontal", () => {
  it("collapses two cells into one with gridSpan=2", async () => {
    const agent = await loadTable([
      ["A", "B", "C"],
      ["D", "E", "F"],
    ]);
    const id = tableBlock(agent.getSnapshot()).id;
    await agent.applyCommand({
      type: "docx:merge-cells-horizontal",
      payload: { tableId: id, row: 0, fromColumn: 0, toColumn: 1 },
    });
    const row = tableBlock(agent.getSnapshot()).rows[0];
    expect(row.cells.length).toBe(2);
    expect(row.cells[0].properties.gridSpan).toBe(2);
  });
});
