// Generate synthetic XLSX fixtures using SheetJS (Apache 2.0).
// Run with `pnpm fixtures:xlsx` (see root package.json).
//
// Each fixture is a self-contained .xlsx file representing a category
// from spec/xlsx/feature-scope.md. They are NOT real-world workbooks —
// they are smoke fixtures that exercise our parser/serializer/handlers.
// Real-world fixtures (Excel-emitted, LibreOffice-emitted) live in
// fixtures/xlsx/real-world/ and are tracked in MANIFEST.md.

import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import * as XLSX from "xlsx";

const here = dirname(fileURLToPath(import.meta.url));
const outRoot = resolve(here, "../fixtures/xlsx/synthetic");

async function write(name, wb) {
  await mkdir(outRoot, { recursive: true });
  const buf = XLSX.write(wb, {
    bookType: "xlsx",
    type: "buffer",
    bookSST: true,
    compression: true,
  });
  const path = resolve(outRoot, `${name}.xlsx`);
  await writeFile(path, buf);
  console.log(`✓ wrote ${path} (${buf.length} bytes)`);
}

// ---------------------------------------------------------------------------
// 01 — single sheet, mixed value kinds
// ---------------------------------------------------------------------------

async function singleSheetNumbers() {
  const wb = XLSX.utils.book_new();
  const data = [
    ["Item", "Qty", "Unit price", "Total"],
    ["Widget", 10, 2.5, 25],
    ["Gadget", 4, 12.99, 51.96],
    ["Sprocket", 100, 0.15, 15],
    ["Sample", true, "n/a", null],
  ];
  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, "Inventory");
  await write("01-single-sheet-numbers", wb);
}

// ---------------------------------------------------------------------------
// 02 — multi-sheet workbook with cross-sheet references
// ---------------------------------------------------------------------------

async function multiSheet() {
  const wb = XLSX.utils.book_new();

  const sales = XLSX.utils.aoa_to_sheet([
    ["Region", "Q1", "Q2", "Q3", "Q4"],
    ["North", 1200, 1500, 1700, 2100],
    ["South", 800, 950, 1100, 1300],
    ["East", 1500, 1700, 1850, 2050],
    ["West", 1100, 1250, 1400, 1700],
  ]);
  XLSX.utils.book_append_sheet(wb, sales, "Sales");

  const expenses = XLSX.utils.aoa_to_sheet([
    ["Region", "Q1", "Q2", "Q3", "Q4"],
    ["North", 700, 850, 950, 1200],
    ["South", 500, 620, 700, 800],
    ["East", 900, 1000, 1100, 1200],
    ["West", 600, 720, 850, 1000],
  ]);
  XLSX.utils.book_append_sheet(wb, expenses, "Expenses");

  const summary = XLSX.utils.aoa_to_sheet([
    ["Region", "Net Q1", "Net Q4", "Total Sales"],
    [
      "North",
      { t: "n", f: "Sales!B2-Expenses!B2" },
      { t: "n", f: "Sales!E2-Expenses!E2" },
      { t: "n", f: "SUM(Sales!B2:E2)" },
    ],
    [
      "South",
      { t: "n", f: "Sales!B3-Expenses!B3" },
      { t: "n", f: "Sales!E3-Expenses!E3" },
      { t: "n", f: "SUM(Sales!B3:E3)" },
    ],
  ]);
  XLSX.utils.book_append_sheet(wb, summary, "Summary");

  await write("02-multi-sheet", wb);
}

// ---------------------------------------------------------------------------
// 03 — formulas: SUM, AVERAGE, IF, VLOOKUP, COUNTIF
// ---------------------------------------------------------------------------

async function formulasBasic() {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["Name", "Score", "Grade", "Pass?"],
    ["Alice", 92, { t: "s", f: "VLOOKUP(B2,Grades!A:B,2,TRUE)" }, { t: "b", f: "B2>=60" }],
    ["Bob", 71, { t: "s", f: "VLOOKUP(B3,Grades!A:B,2,TRUE)" }, { t: "b", f: "B3>=60" }],
    ["Carol", 58, { t: "s", f: "VLOOKUP(B4,Grades!A:B,2,TRUE)" }, { t: "b", f: "B4>=60" }],
    ["Dave", 85, { t: "s", f: "VLOOKUP(B5,Grades!A:B,2,TRUE)" }, { t: "b", f: "B5>=60" }],
    [],
    ["Stats", null, null, null],
    ["Total", { t: "n", f: "SUM(B2:B5)" }, null, null],
    ["Average", { t: "n", f: "AVERAGE(B2:B5)" }, null, null],
    ["Passing", { t: "n", f: "COUNTIF(D2:D5,TRUE)" }, null, null],
    ["Top score", { t: "n", f: "MAX(B2:B5)" }, null, null],
  ]);
  XLSX.utils.book_append_sheet(wb, ws, "Scores");

  const grades = XLSX.utils.aoa_to_sheet([
    [0, "F"],
    [60, "D"],
    [70, "C"],
    [80, "B"],
    [90, "A"],
  ]);
  XLSX.utils.book_append_sheet(wb, grades, "Grades");

  await write("03-formulas-basic", wb);
}

// ---------------------------------------------------------------------------
// 04 — merges + number formats + cell styling
// ---------------------------------------------------------------------------

async function mergedAndFormatted() {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["Quarterly Report", null, null, null],
    ["Region", "Revenue", "Cost", "Margin"],
    ["North", 12500.5, 8200.25, 0.344],
    ["South", 9800.0, 6400.5, 0.347],
    ["East", 15300.75, 9100.0, 0.405],
    ["West", 11200.6, 7800.4, 0.304],
  ]);

  ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];

  ws["A1"].s = { font: { bold: true, sz: 14 }, alignment: { horizontal: "center" } };
  for (const ref of ["A2", "B2", "C2", "D2"]) {
    ws[ref].s = {
      font: { bold: true },
      fill: { fgColor: { rgb: "DDDDDD" } },
      alignment: { horizontal: "center" },
    };
  }
  for (const row of [3, 4, 5, 6]) {
    ws[`B${row}`].z = '"$"#,##0.00';
    ws[`C${row}`].z = '"$"#,##0.00';
    ws[`D${row}`].z = "0.0%";
  }

  ws["!cols"] = [{ wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 12 }];
  ws["!rows"] = [{ hpx: 24 }, { hpx: 18 }];

  XLSX.utils.book_append_sheet(wb, ws, "Q1");
  await write("04-merged-and-formatted", wb);
}

// ---------------------------------------------------------------------------
// 05 — comments + hyperlinks
// ---------------------------------------------------------------------------

async function commentsAndHyperlinks() {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["Resource", "Owner", "Link"],
    ["Spec", "alice", "https://example.com/spec"],
    ["Build log", "bob", "https://example.com/log"],
    ["Repo", "carol", "https://github.com/example/repo"],
  ]);

  ws["A2"].c = [{ a: "alice", t: "Living document, updated weekly." }];
  ws["A3"].c = [{ a: "bob", t: "Generated by CI; do not hand-edit." }];

  ws["C2"].l = { Target: "https://example.com/spec", Tooltip: "Open spec in browser" };
  ws["C3"].l = { Target: "https://example.com/log" };
  ws["C4"].l = { Target: "https://github.com/example/repo" };

  XLSX.utils.book_append_sheet(wb, ws, "Resources");
  await write("05-comments-hyperlinks", wb);
}

// ---------------------------------------------------------------------------
// 06 — large grid for perf smoke
// ---------------------------------------------------------------------------

async function largeGrid() {
  const wb = XLSX.utils.book_new();
  const rows = [["id", "ts", "label", "value", "ratio", "active", "note"]];
  for (let i = 1; i <= 1000; i++) {
    rows.push([
      i,
      new Date(2026, 0, 1, 0, 0, i).toISOString(),
      `row-${i.toString().padStart(4, "0")}`,
      Math.round(Math.sin(i) * 1000) / 10,
      Math.round((i % 100) / 100 / 0.0001) * 0.0001,
      i % 2 === 0,
      `note for row ${i}`,
    ]);
  }
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{ wch: 6 }, { wch: 22 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 24 }];
  XLSX.utils.book_append_sheet(wb, ws, "Data");
  await write("06-large-grid", wb);
}

await singleSheetNumbers();
await multiSheet();
await formulasBasic();
await mergedAndFormatted();
await commentsAndHyperlinks();
await largeGrid();
console.log("\nDone. See fixtures/xlsx/MANIFEST.md.");
