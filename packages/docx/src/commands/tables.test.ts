import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { deterministicIdMinter, ooxml, sha256Hex } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import { serializeDocx } from "../serializer/serialize.js";
import type { BlockNode, Paragraph, Table } from "../model/types.js";
import { paragraphPlainText } from "./helpers.js";
import { DEFAULT_DOC_ROOT_ATTRS, escapeXml, makeSyntheticDocx } from "../test-utils/synthetic.js";

const FIXTURE_PATH = resolve(__dirname, "../../../../fixtures/docx/real-world/04-table-grid.docx");

/**
 * Build a synthetic DOCX whose body is just a `<w:tbl>` and (optionally) a
 * trailing paragraph. `rows` is a 2-D array of cell text; the first row gets
 * `<w:tblHeader/>`. Column widths default to 2000 twips each.
 */
function syntheticTableDoc(
  rows: ReadonlyArray<ReadonlyArray<string>>,
  opts: { columnWidths?: number[]; trailingParagraph?: string } = {}
): string {
  const cols = rows[0]?.length ?? 0;
  const widths = opts.columnWidths ?? new Array(cols).fill(2000);
  const grid = widths.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  const trXml = rows
    .map((row, r) => {
      const cells = row
        .map(
          (text, c) =>
            `<w:tc><w:tcPr><w:tcW w:w="${widths[c]}" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:tc>`
        )
        .join("");
      const trPr = r === 0 ? `<w:trPr><w:tblHeader/></w:trPr>` : "";
      return `<w:tr>${trPr}${cells}</w:tr>`;
    })
    .join("");
  const tblPr = `<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>`;
  const tableXml = `<w:tbl>${tblPr}<w:tblGrid>${grid}</w:tblGrid>${trXml}</w:tbl>`;
  const trailing = opts.trailingParagraph
    ? `<w:p><w:r><w:t xml:space="preserve">${escapeXml(opts.trailingParagraph)}</w:t></w:r></w:p>`
    : `<w:p/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>${tableXml}${trailing}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
}

async function loadSyntheticTable(
  rows: ReadonlyArray<ReadonlyArray<string>>,
  opts: { columnWidths?: number[]; trailingParagraph?: string } = {}
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

function plainTextOf(block: BlockNode): string {
  if (block.kind !== "paragraph") return "";
  return paragraphPlainText(block);
}

describe("docx tables — parser + commands (P1.3 / W7)", () => {
  it("parses a synthetic 2×3 table with a heading row + body rows", async () => {
    const agent = await loadSyntheticTable(
      [
        ["A", "B", "C"],
        ["1", "2", "3"],
      ],
      { columnWidths: [1000, 2000, 3000] }
    );
    const tbl = tableBlock(agent.getSnapshot());
    expect(tbl.kind).toBe("table");
    expect(tbl.grid).toEqual([{ w: 1000 }, { w: 2000 }, { w: 3000 }]);
    expect(tbl.rows.length).toBe(2);
    expect(tbl.rows[0].properties.header).toBe(true);
    expect(tbl.rows[1].properties.header).toBeUndefined();
    expect(tbl.rows[0].cells.map((c) => plainTextOf(c.body[0]))).toEqual(["A", "B", "C"]);
    expect(tbl.rows[1].cells.map((c) => plainTextOf(c.body[0]))).toEqual(["1", "2", "3"]);
    // Untouched table → raw cache present.
    expect(tbl.raw).toBeDefined();
  });

  it("parses 04-table-grid.docx with the right shape and at least one cell text", async () => {
    const buf = await readFile(FIXTURE_PATH);
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const tbl = tableBlock(snap);
    expect(tbl.grid.length).toBe(3);
    expect(tbl.rows.length).toBe(4);
    // First row of the fixture is a header row.
    expect(tbl.rows[0].properties.header).toBe(true);
    // Some cell in the table has visible text.
    let total = "";
    for (const row of tbl.rows) {
      for (const cell of row.cells) {
        for (const block of cell.body) {
          total += plainTextOf(block);
        }
      }
    }
    expect(total.length).toBeGreaterThan(0);
  });

  it("byte-preservation: untouched parse → serialize keeps word/document.xml SHA-256 unchanged", async () => {
    const buf = await readFile(FIXTURE_PATH);
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const out = await serializeDocx(snap);
    const reloaded = await ooxml.OoxmlContainer.load(out);
    const beforeHash = sha256Hex(snap.container.readBytes("word/document.xml"));
    const afterHash = sha256Hex(reloaded.readBytes("word/document.xml"));
    expect(afterHash).toBe(beforeHash);
  });

  it("insert-table: appends and survives parse(serialize(s))", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:p><w:r><w:t xml:space="preserve">hello</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
    });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const before = agent.getSnapshot().root.body.length;
    const m = await agent.applyCommand({
      type: "docx:insert-table",
      payload: { at: { paragraph: before }, rows: 2, cols: 3, columnWidths: [1500, 2500, 3000] },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const next = agent.getSnapshot();
    expect(next.root.body.length).toBe(before + 1);
    const inserted = tableBlock(next);
    expect(inserted.grid).toEqual([{ w: 1500 }, { w: 2500 }, { w: 3000 }]);
    expect(inserted.rows.length).toBe(2);
    expect(inserted.rows[0].cells.length).toBe(3);
    expect(inserted.raw).toBeUndefined();

    // Round-trip through parse(serialize(...)).
    const out = await agent.exportFile();
    const reparsed = await parseDocx(out, { idMinter: deterministicIdMinter("rt") });
    const reparsedTbl = tableBlock(reparsed);
    expect(reparsedTbl.rows.length).toBe(2);
    expect(reparsedTbl.grid.length).toBe(3);
    expect(reparsedTbl.grid.map((g) => g.w)).toEqual([1500, 2500, 3000]);
  });

  it("insert-table: rejects rows/cols mismatch with columnWidths.length", async () => {
    const agent = await loadSyntheticTable([["x"]]);
    const m = await agent.applyCommand({
      type: "docx:insert-table",
      payload: {
        at: { paragraph: 0 },
        rows: 2,
        cols: 3,
        columnWidths: [1000, 2000], // wrong length
      },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("set-cell-content: replaces a cell, marks dirty, untouched table parts stay raw-cached", async () => {
    const agent = await loadSyntheticTable([
      ["A", "B"],
      ["1", "2"],
    ]);
    const before = tableBlock(agent.getSnapshot());
    const targetCellId = before.rows[1].cells[0].id;
    const newPara: Paragraph = {
      kind: "paragraph",
      id: "replaced-p",
      properties: {},
      children: [
        {
          kind: "run",
          id: "replaced-r",
          properties: {},
          children: [{ kind: "text", id: "replaced-t", text: "REPLACED", xmlSpacePreserve: false }],
        },
      ],
    };
    const m = await agent.applyCommand({
      type: "docx:set-cell-content",
      payload: {
        tableId: before.id,
        row: 1,
        col: 0,
        content: [newPara],
      },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const after = tableBlock(agent.getSnapshot());
    // After mutation, raw must be dropped so the serializer regenerates the
    // table from the typed model.
    expect(after.raw).toBeUndefined();
    expect(after.rows[1].cells[0].id).toBe(targetCellId);
    expect(plainTextOf(after.rows[1].cells[0].body[0])).toBe("REPLACED");
    // Untouched cell content survives intact.
    expect(plainTextOf(after.rows[0].cells[0].body[0])).toBe("A");
    expect(plainTextOf(after.rows[1].cells[1].body[0])).toBe("2");
    // Body is dirty.
    expect(agent.getSnapshot().dirty.body).toBe(true);
  });

  it("set-cell-content: unknown tableId → unknown-target", async () => {
    const agent = await loadSyntheticTable([["x"]]);
    const m = await agent.applyCommand({
      type: "docx:set-cell-content",
      payload: { tableId: "no-such-table", row: 0, col: 0, content: [] },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });

  it("set-cell-content: out-of-bounds row/col → unknown-target", async () => {
    const agent = await loadSyntheticTable([
      ["A", "B"],
      ["1", "2"],
    ]);
    const tbl = tableBlock(agent.getSnapshot());
    const m1 = await agent.applyCommand({
      type: "docx:set-cell-content",
      payload: { tableId: tbl.id, row: 99, col: 0, content: [] },
      source: "human",
    });
    expect(m1.status).toBe("rejected");
    expect(m1.rejection?.code).toBe("unknown-target");
    const m2 = await agent.applyCommand({
      type: "docx:set-cell-content",
      payload: { tableId: tbl.id, row: 0, col: 99, content: [] },
      source: "human",
    });
    expect(m2.status).toBe("rejected");
    expect(m2.rejection?.code).toBe("unknown-target");
  });

  it("insert-row: appends and inherits gridCol widths", async () => {
    const agent = await loadSyntheticTable(
      [
        ["A", "B", "C"],
        ["1", "2", "3"],
      ],
      { columnWidths: [1000, 2000, 3000] }
    );
    const tbl0 = tableBlock(agent.getSnapshot());
    const m = await agent.applyCommand({
      type: "docx:insert-row",
      payload: { tableId: tbl0.id, at: tbl0.rows.length },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const tbl = tableBlock(agent.getSnapshot());
    expect(tbl.rows.length).toBe(3);
    const newRow = tbl.rows[2];
    expect(newRow.cells.length).toBe(3);
    expect(newRow.cells.map((c) => c.properties.tcW?.value)).toEqual([1000, 2000, 3000]);
    // Existing header flag stays only on the first declared row.
    expect(tbl.rows[0].properties.header).toBe(true);
    expect(newRow.properties.header).toBeUndefined();
  });

  it("insert-row at index 0: header flag stays only on the originally first row", async () => {
    const agent = await loadSyntheticTable([
      ["H1", "H2"],
      ["b1", "b2"],
    ]);
    const tbl0 = tableBlock(agent.getSnapshot());
    const m = await agent.applyCommand({
      type: "docx:insert-row",
      payload: { tableId: tbl0.id, at: 0 },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const tbl = tableBlock(agent.getSnapshot());
    expect(tbl.rows.length).toBe(3);
    // Newly inserted row has no header flag; the originally-declared header
    // row (now at index 1) keeps it.
    expect(tbl.rows[0].properties.header).toBeUndefined();
    expect(tbl.rows[1].properties.header).toBe(true);
    expect(tbl.rows[2].properties.header).toBeUndefined();
  });

  it("insert-column: adds a gridCol + tc to every row, widths sum stays consistent", async () => {
    const agent = await loadSyntheticTable(
      [
        ["A", "B"],
        ["1", "2"],
      ],
      { columnWidths: [2000, 2000] }
    );
    const tbl0 = tableBlock(agent.getSnapshot());
    const widthsBefore = tbl0.grid.reduce((sum, g) => sum + (g.w ?? 0), 0);
    const m = await agent.applyCommand({
      type: "docx:insert-column",
      payload: { tableId: tbl0.id, at: 1, width: 1500 },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const tbl = tableBlock(agent.getSnapshot());
    expect(tbl.grid.length).toBe(3);
    expect(tbl.grid.map((g) => g.w)).toEqual([2000, 1500, 2000]);
    for (const row of tbl.rows) {
      expect(row.cells.length).toBe(3);
      expect(row.cells[1].properties.tcW?.value).toBe(1500);
    }
    const widthsAfter = tbl.grid.reduce((sum, g) => sum + (g.w ?? 0), 0);
    expect(widthsAfter).toBe(widthsBefore + 1500);
  });

  it("insert-column at the right edge (at === cols) appends", async () => {
    const agent = await loadSyntheticTable(
      [
        ["A", "B"],
        ["1", "2"],
      ],
      { columnWidths: [1000, 1000] }
    );
    const tbl0 = tableBlock(agent.getSnapshot());
    const cols = tbl0.grid.length;
    const m = await agent.applyCommand({
      type: "docx:insert-column",
      payload: { tableId: tbl0.id, at: cols, width: 500 },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const tbl = tableBlock(agent.getSnapshot());
    expect(tbl.grid.length).toBe(cols + 1);
    expect(tbl.grid[cols].w).toBe(500);
    for (const row of tbl.rows) {
      expect(row.cells.length).toBe(cols + 1);
      const newCell = row.cells[cols];
      expect(newCell.properties.tcW?.value).toBe(500);
      // New cell carries one empty paragraph by default.
      expect(newCell.body.length).toBe(1);
      expect(newCell.body[0].kind).toBe("paragraph");
    }
  });

  it("nested table: parses and round-trips structurally", async () => {
    // Outer 1x1 table whose only cell contains a nested 1x2 table.
    const inner = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="500"/><w:gridCol w:w="500"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t xml:space="preserve">y</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const outer = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${inner}<w:p/></w:tc></w:tr></w:tbl>`;
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>${outer}<w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
    const buf = await makeSyntheticDocx({ documentXml: docXml });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const outerTbl = tableBlock(agent.getSnapshot());
    expect(outerTbl.rows.length).toBe(1);
    const cellBody = outerTbl.rows[0].cells[0].body;
    const nested = cellBody.find((b): b is Table => b.kind === "table");
    expect(nested).toBeDefined();
    expect(nested?.grid.length).toBe(2);
    // Mutate the nested table: insert a row.
    const mutNested = nested!;
    const m = await agent.applyCommand({
      type: "docx:insert-row",
      payload: { tableId: mutNested.id, at: mutNested.rows.length },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const after = agent.getSnapshot();
    const outerAfter = tableBlock(after);
    // Outer table's `raw` cache is dropped because a descendant changed.
    expect(outerAfter.raw).toBeUndefined();
    const nestedAfter = outerAfter.rows[0].cells[0].body.find((b): b is Table => b.kind === "table");
    expect(nestedAfter?.rows.length).toBe(2);
  });

  it("set-cell-content: rejects content containing a table whose id matches the target table (cycle)", async () => {
    const agent = await loadSyntheticTable([
      ["A", "B"],
      ["1", "2"],
    ]);
    const tbl = tableBlock(agent.getSnapshot());
    // Build a "table" payload with the same id as the target.
    const cyclic: Table = {
      kind: "table",
      id: tbl.id,
      properties: {},
      grid: [{ w: 100 }],
      rows: [
        {
          kind: "table-row",
          id: "cyclic-row",
          properties: {},
          cells: [
            {
              kind: "table-cell",
              id: "cyclic-cell",
              properties: {},
              body: [
                {
                  kind: "paragraph",
                  id: "cyclic-p",
                  properties: {},
                  children: [{ kind: "run", id: "cyclic-r", properties: {}, children: [] }],
                },
              ],
            },
          ],
        },
      ],
    };
    const m = await agent.applyCommand({
      type: "docx:set-cell-content",
      payload: { tableId: tbl.id, row: 0, col: 0, content: [cyclic] },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });

  it("preserves gridSpan and vMerge on round-trip; rejects writes into a vMerge continuation cell", async () => {
    // 2 rows × 2 cols, with the (0,0) cell carrying gridSpan=2 (so row 0
    // visually has only one wide cell), and (1,0) is a vMerge restart that
    // (1,1) "continues" — synthetic merge layout.
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">wide</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">left-restart</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
    const buf = await makeSyntheticDocx({ documentXml: docXml });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const tbl = tableBlock(agent.getSnapshot());
    expect(tbl.rows[0].cells[0].properties.gridSpan).toBe(2);
    expect(tbl.rows[1].cells[0].properties.vMerge).toBe("restart");
    expect(tbl.rows[1].cells[1].properties.vMerge).toBe("continue");

    // Writing into the vMerge continuation cell must be rejected.
    const m = await agent.applyCommand({
      type: "docx:set-cell-content",
      payload: { tableId: tbl.id, row: 1, col: 1, content: [] },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("merged-cell-not-supported");
  });
});
