#!/usr/bin/env node
/**
 * Diagnostic for the masterthesis fixture: dumps every table's typed
 * properties (width, indent, justification, layout, borders) and the
 * raw XML of any `<w:tblPr>` so we can see why the quadrant tables
 * render off-center in the editor.
 */
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { DocxAgent } from "@officeai/docx";

const FIXTURE = resolve(
  process.cwd(),
  "..",
  "Masterthesis_Rohfassung_Verification_Asymmetry Kopie.docx"
);

const buf = await readFile(FIXTURE);
const agent = await DocxAgent.fromBuffer(buf);
const snap = agent.getSnapshot();
const body = snap.root.body;

let n = 0;
for (let i = 0; i < body.length; i++) {
  const b = body[i];
  if (b.kind !== "table") continue;
  n++;
  const sample = b.rows[0]?.cells[0]?.body
    ?.map((bb) =>
      bb.kind === "paragraph"
        ? bb.children
            .flatMap((inl) =>
              inl.kind === "run"
                ? inl.children.flatMap((c) => (c.kind === "text" ? c.text : ""))
                : []
            )
            .join("")
        : ""
    )
    .join(" ")
    .slice(0, 60);
  console.log(`\n=== Block ${i}: table #${n} (rows=${b.rows.length}) cell0,0="${sample}" ===`);
  console.log("table.properties:", JSON.stringify(b.properties ?? null, null, 2));
  console.log("table.grid:", JSON.stringify(b.grid ?? null, null, 2));
  console.log("table top-level keys:", Object.keys(b));
  // First row cell properties
  const cells = b.rows[0]?.cells.map((c) => ({
    keys: Object.keys(c),
    properties: c.properties,
  }));
  console.log("row0 cells:", JSON.stringify(cells, null, 2));
}
console.log(`\nTotal tables: ${n}`);
