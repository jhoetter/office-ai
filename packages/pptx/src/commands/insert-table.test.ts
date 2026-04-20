import { execFile } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { join, resolve } from "node:path";
import { promisify } from "node:util";
import { describe, expect, it } from "vitest";
import { PptxAgent } from "../agent/agent.js";
import type { TableShape } from "../model/types.js";

const exec = promisify(execFile);
const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

/**
 * Resolve a usable LibreOffice executable. Mirrors the helper in
 * `packages/xlsx/src/serializer/charts.libreoffice.test.ts` so dev
 * workstations without LibreOffice still see a green suite (the
 * `describeIfLO` tests are skipped, not failed).
 */
function findSoffice(): string | null {
  const candidates = [
    process.env.SOFFICE,
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/opt/homebrew/bin/soffice",
  ].filter((p): p is string => typeof p === "string" && p.length > 0);
  return candidates.find((p) => existsSync(p)) ?? null;
}

const SOFFICE = findSoffice();

async function loadBlankAgent(): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, "01-blank.pptx"));
  return PptxAgent.fromBuffer(buf);
}

function flattenCellText(cell: TableShape["rows"][number]["cells"][number]): string {
  return cell.txBody.paragraphs
    .flatMap((p) => p.runs.filter((r) => !r.isLineBreak).map((r) => r.text))
    .join("");
}

function findFirstTable(agent: PptxAgent): TableShape | null {
  for (const slide of agent.getSnapshot().root.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind === "table") return shape;
    }
  }
  return null;
}

describe("D1: pptx:insert-table", () => {
  it("rejects non-positive rows / cols", async () => {
    const agent = await loadBlankAgent();
    const r1 = await agent.applyCommand({
      type: "pptx:insert-table",
      payload: { slideIndex: 0, rows: 0, cols: 2 },
    });
    expect(r1.rejection?.code).toBe("invalid-payload");
    const r2 = await agent.applyCommand({
      type: "pptx:insert-table",
      payload: { slideIndex: 0, rows: 2, cols: -1 },
    });
    expect(r2.rejection?.code).toBe("invalid-payload");
  });

  it("rejects unknown slide index", async () => {
    const agent = await loadBlankAgent();
    const m = await agent.applyCommand({
      type: "pptx:insert-table",
      payload: { slideIndex: 99, rows: 2, cols: 2 },
    });
    expect(m.rejection?.code).toBe("unknown-target");
  });

  it("inserts a typed TableShape with the requested geometry and seeded cells", async () => {
    const agent = await loadBlankAgent();
    const before = agent.getSnapshot().root.slides[0]!.shapes.length;
    await agent.applyCommand({
      type: "pptx:insert-table",
      payload: {
        slideIndex: 0,
        rows: 3,
        cols: 4,
        cells: [
          ["Q1", "Q2", "Q3", "Q4"],
          ["1", "2", "3", "4"],
        ],
      },
    });
    const slide = agent.getSnapshot().root.slides[0]!;
    expect(slide.shapes.length).toBe(before + 1);
    const table = findFirstTable(agent)!;
    expect(table).toBeDefined();
    expect(table.kind).toBe("table");
    expect(table.rows.length).toBe(3);
    expect(table.columnWidths.length).toBe(4);
    expect(table.rows.every((r) => r.cells.length === 4)).toBe(true);
    expect(flattenCellText(table.rows[0]!.cells[0]!)).toBe("Q1");
    expect(flattenCellText(table.rows[0]!.cells[3]!)).toBe("Q4");
    expect(flattenCellText(table.rows[1]!.cells[2]!)).toBe("3");
    expect(flattenCellText(table.rows[2]!.cells[0]!)).toBe("");
    expect(table.graphicDataUri).toBe(
      "http://schemas.openxmlformats.org/drawingml/2006/table"
    );
    // Cell text picked up the master/fallback default font so PowerPoint
    // doesn't fall back to its built-in 24pt body style.
    const seedRun = table.rows[0]!.cells[0]!.txBody.paragraphs[0]!.runs[0]!;
    expect(seedRun.properties.fontFamily).toBeTruthy();
    expect(seedRun.properties.fontSizeHundredths).toBeGreaterThan(0);
    expect(agent.getSnapshot().dirty.slides.size).toBe(1);
  });

  it("survives parse → insert → serialize → parse with the table intact", async () => {
    const agent = await loadBlankAgent();
    await agent.applyCommand({
      type: "pptx:insert-table",
      payload: {
        slideIndex: 0,
        rows: 2,
        cols: 3,
        cells: [
          ["A", "B", "C"],
          ["D", "E", "F"],
        ],
      },
    });
    const out = await agent.exportFile();
    const reloaded = await PptxAgent.fromBuffer(out);
    const reTable = findFirstTable(reloaded)!;
    expect(reTable).toBeDefined();
    expect(reTable.rows.length).toBe(2);
    expect(reTable.columnWidths.length).toBe(3);
    expect(flattenCellText(reTable.rows[0]!.cells[0]!)).toBe("A");
    expect(flattenCellText(reTable.rows[1]!.cells[2]!)).toBe("F");
  });

  it("centres the table on the slide when xEmu/yEmu are omitted", async () => {
    const agent = await loadBlankAgent();
    await agent.applyCommand({
      type: "pptx:insert-table",
      payload: { slideIndex: 0, rows: 2, cols: 2, widthEmu: 4_000_000, heightEmu: 2_000_000 },
    });
    const table = findFirstTable(agent)!;
    const slideSize = agent.getSnapshot().root.slideSize;
    expect(table.position!.xEmu).toBe(Math.round((slideSize.cxEmu - 4_000_000) / 2));
    expect(table.position!.yEmu).toBe(Math.round((slideSize.cyEmu - 2_000_000) / 2));
    expect(table.size!.cxEmu).toBe(4_000_000);
    expect(table.size!.cyEmu).toBe(2_000_000);
    // Column widths and row heights should sum exactly to the frame extent.
    const widthSum = table.columnWidths.reduce((a, b) => a + b, 0);
    const heightSum = table.rows.reduce((a, r) => a + r.height, 0);
    expect(widthSum).toBe(4_000_000);
    expect(heightSum).toBe(2_000_000);
  });
});

const describeIfLO = SOFFICE ? describe : describe.skip;

describeIfLO("D1: pptx:insert-table — LibreOffice headless round-trip", () => {
  it(
    "the inserted table survives a LibreOffice convert-to pptx pass",
    async () => {
      const agent = await loadBlankAgent();
      await agent.applyCommand({
        type: "pptx:insert-table",
        payload: {
          slideIndex: 0,
          rows: 3,
          cols: 2,
          cells: [
            ["Header A", "Header B"],
            ["row1a", "row1b"],
            ["row2a", "row2b"],
          ],
        },
      });
      const out = await agent.exportFile();

      const tmpDir = resolve("/tmp/officeai-pptx-insert-table");
      await rm(tmpDir, { recursive: true, force: true });
      await mkdir(tmpDir, { recursive: true });
      const inPath = resolve(tmpDir, "with-table.pptx");
      const outDir = resolve(tmpDir, "converted");
      await mkdir(outDir, { recursive: true });
      await writeFile(inPath, new Uint8Array(out));

      const profile = `-env:UserInstallation=file://${tmpDir}/lo-profile`;
      const { stdout, stderr } = await exec(SOFFICE!, [
        profile,
        "--headless",
        "--convert-to",
        "pptx",
        inPath,
        "--outdir",
        outDir,
      ]);
      void stdout;
      void stderr;

      const reEmitted = new Uint8Array(await readFile(resolve(outDir, "with-table.pptx")));
      const reloaded = await PptxAgent.fromBuffer(reEmitted);
      const reTable = findFirstTable(reloaded);
      expect(reTable).toBeDefined();
      expect(reTable!.rows.length).toBe(3);
      expect(reTable!.columnWidths.length).toBe(2);
      // LO sometimes rewrites cell text formatting; we just verify the
      // values made it through, not the per-run properties.
      expect(flattenCellText(reTable!.rows[0]!.cells[0]!)).toBe("Header A");
      expect(flattenCellText(reTable!.rows[2]!.cells[1]!)).toBe("row2b");
    },
    60_000
  );
});
