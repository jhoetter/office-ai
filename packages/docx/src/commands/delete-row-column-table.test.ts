import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { DEFAULT_DOC_ROOT_ATTRS, escapeXml, makeSyntheticDocx } from "../test-utils/synthetic.js";
import type { BlockNode, Table } from "../model/types.js";

function syntheticTableDoc(
  rows: ReadonlyArray<ReadonlyArray<string>>,
  opts: { columnWidths?: number[] } = {}
): string {
  const cols = rows[0]?.length ?? 0;
  const widths = opts.columnWidths ?? new Array(cols).fill(2000);
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

async function loadSyntheticTable(
  rows: ReadonlyArray<ReadonlyArray<string>>,
  opts: { columnWidths?: number[] } = {}
): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: syntheticTableDoc(rows, opts) });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function tableBlock(snap: { root: { body: ReadonlyArray<BlockNode> } }, index = 0): Table {
  for (const b of snap.root.body) {
    if (b.kind === "table") {
      if (index === 0) return b;
      index--;
    }
  }
  throw new Error("no table found in snapshot body");
}

describe("docx delete-row / delete-column / delete-table", () => {
  it("delete-row removes the targeted row and dirty-flags the body", async () => {
    const agent = await loadSyntheticTable([
      ["A", "B"],
      ["1", "2"],
      ["x", "y"],
    ]);
    const tbl0 = tableBlock(agent.getSnapshot());
    const m = await agent.applyCommand({
      type: "docx:delete-row",
      payload: { tableId: tbl0.id, at: 1 },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const tbl = tableBlock(agent.getSnapshot());
    expect(tbl.rows.length).toBe(2);
    const texts = tbl.rows.map((r) =>
      r.cells.map((c) => {
        const p = c.body[0];
        if (p.kind !== "paragraph") return "";
        let out = "";
        for (const child of p.children) {
          if (child.kind === "run") for (const t of child.children) if (t.kind === "text") out += t.text;
        }
        return out;
      })
    );
    expect(texts).toEqual([
      ["A", "B"],
      ["x", "y"],
    ]);
    expect(agent.getSnapshot().dirty.body).toBe(true);
  });

  it("delete-row rejects out-of-range indices", async () => {
    const agent = await loadSyntheticTable([["A"], ["1"]]);
    const tbl0 = tableBlock(agent.getSnapshot());
    const m = await agent.applyCommand({
      type: "docx:delete-row",
      payload: { tableId: tbl0.id, at: 5 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-position");
  });

  it("delete-row refuses to remove the last surviving row", async () => {
    const agent = await loadSyntheticTable([["only"]]);
    const tbl0 = tableBlock(agent.getSnapshot());
    const m = await agent.applyCommand({
      type: "docx:delete-row",
      payload: { tableId: tbl0.id, at: 0 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-position");
  });

  it("delete-column drops a gridCol entry and the matching tc on every row", async () => {
    const agent = await loadSyntheticTable(
      [
        ["A", "B", "C"],
        ["1", "2", "3"],
      ],
      { columnWidths: [1000, 2000, 3000] }
    );
    const tbl0 = tableBlock(agent.getSnapshot());
    const m = await agent.applyCommand({
      type: "docx:delete-column",
      payload: { tableId: tbl0.id, at: 1 },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const tbl = tableBlock(agent.getSnapshot());
    expect(tbl.grid.length).toBe(2);
    expect(tbl.grid.map((g) => g.w)).toEqual([1000, 3000]);
    for (const row of tbl.rows) {
      expect(row.cells.length).toBe(2);
      expect(row.cells[0].properties.tcW?.value).toBe(1000);
      expect(row.cells[1].properties.tcW?.value).toBe(3000);
    }
  });

  it("delete-column rejects when the table has only one column", async () => {
    const agent = await loadSyntheticTable([["A"], ["1"]]);
    const tbl0 = tableBlock(agent.getSnapshot());
    const m = await agent.applyCommand({
      type: "docx:delete-column",
      payload: { tableId: tbl0.id, at: 0 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-position");
  });

  it("delete-table removes the top-level table block", async () => {
    const agent = await loadSyntheticTable([
      ["A", "B"],
      ["1", "2"],
    ]);
    const before = agent.getSnapshot();
    const tbl0 = tableBlock(before);
    const m = await agent.applyCommand({
      type: "docx:delete-table",
      payload: { tableId: tbl0.id },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const after = agent.getSnapshot();
    const stillHasTable = after.root.body.some((b) => b.kind === "table");
    expect(stillHasTable).toBe(false);
    expect(after.dirty.body).toBe(true);
  });

  it("delete-table rejects on unknown id", async () => {
    const agent = await loadSyntheticTable([["A"], ["1"]]);
    const m = await agent.applyCommand({
      type: "docx:delete-table",
      payload: { tableId: "no-such-table" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });
});
