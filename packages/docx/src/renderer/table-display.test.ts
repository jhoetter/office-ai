import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { docToPM } from "./doc-to-pm.js";
import { docxSchema } from "./schema.js";
import { DEFAULT_DOC_ROOT_ATTRS, makeSyntheticDocx } from "../test-utils/synthetic.js";

/**
 * The typed `Table` model (P1.3 / W7) carries full row + cell + cell-body
 * structure; before this commit the renderer atomized every table to a
 * `[table]` chip regardless. These tests verify that the projection
 * surfaces the typed cell text and respects gridSpan, header rows, and
 * vMerge continuation skip — without disturbing top-level paragraph
 * indexing (cells stay read-only for now).
 */

async function loadAgentWithBody(bodyXml: string): Promise<DocxAgent> {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>${bodyXml}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
  const buf = await makeSyntheticDocx({ documentXml: xml });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function tableSpec(node: import("prosemirror-model").Node): unknown[] {
  const spec = docxSchema.nodes.table.spec;
  if (typeof spec.toDOM !== "function") throw new Error("table spec has no toDOM");
  return spec.toDOM(node) as unknown as unknown[];
}

/**
 * Phase 2 of the renderer wraps every cell run in a `<span class="pm-table-run">`
 * so that bold/italic/font/highlight marks survive into the DOM. The
 * tests that originally asserted `p[2] === "raw text"` need to walk
 * through the span layer to recover the literal text — that is what
 * this helper does. It joins multiple runs into a single string so the
 * assertions stay readable.
 */
function cellTextOf(cell: unknown): string {
  const p = findElement(cell, "p");
  if (!Array.isArray(p)) return "";
  let out = "";
  for (let i = 2; i < p.length; i++) {
    const child = p[i];
    if (typeof child === "string") {
      out += child;
    } else if (Array.isArray(child) && child[0] === "span") {
      // Span structure: ["span", attrs, textOrNestedSpec]
      const inner = child[2];
      if (typeof inner === "string") out += inner;
      else if (Array.isArray(inner)) out += extractInnerText(inner);
    }
  }
  return out;
}

function extractInnerText(spec: unknown[]): string {
  // Walk through nested mark wrappers (strong/em/u/s) until we hit text.
  let cur: unknown = spec;
  while (Array.isArray(cur)) {
    const last = cur[cur.length - 1];
    if (typeof last === "string") return last;
    cur = last;
  }
  return "";
}

function findElement(spec: unknown, tag: string): unknown[] | null {
  if (!Array.isArray(spec)) return null;
  if (spec[0] === tag) return spec;
  for (const child of spec) {
    if (Array.isArray(child)) {
      const found = findElement(child, tag);
      if (found) return found;
    }
  }
  return null;
}

function collectAllElements(spec: unknown, tag: string, out: unknown[][] = []): unknown[][] {
  if (!Array.isArray(spec)) return out;
  if (spec[0] === tag) out.push(spec);
  for (const child of spec) {
    if (Array.isArray(child)) collectAllElements(child, tag, out);
  }
  return out;
}

describe("renderer typed-table display", () => {
  it("projects a 2x2 table into a structured <table> DOM with cell text", async () => {
    const tbl = `
      <w:tbl>
        <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
        <w:tr>
          <w:tc><w:p><w:r><w:t xml:space="preserve">A1</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t xml:space="preserve">A2</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t xml:space="preserve">B1</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t xml:space="preserve">B2</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    `;
    const agent = await loadAgentWithBody(tbl);
    const pm = docToPM(agent.getSnapshot());
    expect(pm.child(0).type.name).toBe("table");
    const spec = tableSpec(pm.child(0));
    const tableEl = findElement(spec, "table");
    expect(tableEl).not.toBeNull();
    const cells = collectAllElements(spec, "td");
    expect(cells).toHaveLength(4);
    const cellText = cells.map(cellTextOf);
    expect(cellText).toEqual(["A1", "A2", "B1", "B2"]);
  });

  it("renders a header row using <th> instead of <td>", async () => {
    const tbl = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2500"/></w:tblGrid>
        <w:tr>
          <w:trPr><w:tblHeader/></w:trPr>
          <w:tc><w:p><w:r><w:t xml:space="preserve">Heading</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    `;
    const agent = await loadAgentWithBody(tbl);
    const pm = docToPM(agent.getSnapshot());
    const spec = tableSpec(pm.child(0));
    expect(collectAllElements(spec, "th")).toHaveLength(1);
    expect(collectAllElements(spec, "td")).toHaveLength(1);
  });

  it("emits colspan for cells with gridSpan > 1", async () => {
    const tbl = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
        <w:tr>
          <w:tc>
            <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
            <w:p><w:r><w:t xml:space="preserve">spans both columns</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
    `;
    const agent = await loadAgentWithBody(tbl);
    const pm = docToPM(agent.getSnapshot());
    const spec = tableSpec(pm.child(0));
    const tds = collectAllElements(spec, "td");
    expect(tds).toHaveLength(1);
    const attrs = tds[0]![1] as Record<string, string>;
    expect(attrs.colspan).toBe("2");
  });

  it("skips vMerge=continue cells so the merged cell renders once", async () => {
    const tbl = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2500"/></w:tblGrid>
        <w:tr>
          <w:tc>
            <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
            <w:p><w:r><w:t xml:space="preserve">spans down</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
        <w:tr>
          <w:tc>
            <w:tcPr><w:vMerge/></w:tcPr>
            <w:p/>
          </w:tc>
        </w:tr>
      </w:tbl>
    `;
    const agent = await loadAgentWithBody(tbl);
    const pm = docToPM(agent.getSnapshot());
    const spec = tableSpec(pm.child(0));
    expect(collectAllElements(spec, "td")).toHaveLength(1);
  });

  it("falls back to the [table] placeholder for tables with no projected rows", async () => {
    const node = docxSchema.nodes.table.create({
      tableId: "t1",
      rawJson: "null",
      tableJson: JSON.stringify({ rows: [] }),
    });
    const spec = docxSchema.nodes.table.spec.toDOM!(node) as unknown as unknown[];
    expect(spec[0]).toBe("div");
    expect(spec[2]).toBe("[table]");
  });

  it("does not introduce extra top-level PM blocks for a table", async () => {
    const tbl = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2500"/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
    `;
    const agent = await loadAgentWithBody(tbl);
    const pm = docToPM(agent.getSnapshot());
    expect(pm.childCount).toBe(2);
    expect(pm.child(0).type.name).toBe("table");
    expect(pm.child(1).type.name).toBe("section_break");
  });
});
