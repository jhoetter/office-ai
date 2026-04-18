import { readFile } from "node:fs/promises";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { parseXlsx } from "./parse.js";
import { resolveTargetPath } from "./parse.js";
import { XlsxParseError } from "./errors.js";

const here = dirname(fileURLToPath(import.meta.url));
const fixtures = resolve(here, "../../../../fixtures/xlsx/synthetic");

async function loadFixture(name: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(resolve(fixtures, name)));
}

describe("parseXlsx", () => {
  it("parses 01-single-sheet-numbers and exposes a typed sheet", async () => {
    const buf = await loadFixture("01-single-sheet-numbers.xlsx");
    const snap = await parseXlsx(buf, { idMinter: deterministicIdMinter("n") });

    expect(snap.format).toBe("xlsx");
    expect(snap.revision).toBe(0);
    expect(snap.root.sheets).toHaveLength(1);
    expect(snap.root.sheets[0].name).toBe("Inventory");
    expect(snap.root.sheets[0].kind).toBe("worksheet");
    expect(snap.root.sheets[0].partPath).toBe("xl/worksheets/sheet1.xml");
    expect(snap.root.sheets[0].state).toBe("visible");
    expect(snap.root.sheets[0].index).toBe(0);
  });

  it("parses 02-multi-sheet and preserves tab order", async () => {
    const buf = await loadFixture("02-multi-sheet.xlsx");
    const snap = await parseXlsx(buf);
    expect(snap.root.sheets.map((s) => s.name)).toEqual(["Sales", "Expenses", "Summary"]);
    snap.root.sheets.forEach((s, i) => {
      expect(s.index).toBe(i);
      expect(s.partPath.startsWith("xl/worksheets/sheet")).toBe(true);
    });
  });

  it("computes part hashes for every zip entry", async () => {
    const buf = await loadFixture("03-formulas-basic.xlsx");
    const snap = await parseXlsx(buf);
    expect(Object.keys(snap.partHashes).length).toBeGreaterThan(5);
    expect(snap.partHashes["xl/workbook.xml"]).toMatch(/^[0-9a-f]{64}$/);
    expect(snap.partHashes["xl/worksheets/sheet1.xml"]).toMatch(/^[0-9a-f]{64}$/);
    for (const [path, hash] of Object.entries(snap.partHashes)) {
      expect(hash, `hash for ${path}`).toMatch(/^[0-9a-f]{64}$/);
    }
  });

  it("classifies non-modeled parts (theme, app props, metadata) as opaque", async () => {
    const buf = await loadFixture("06-large-grid.xlsx");
    const snap = await parseXlsx(buf);
    const opaquePaths = [...snap.root.opaqueParts.keys()].sort();
    expect(opaquePaths).toContain("xl/theme/theme1.xml");
    expect(opaquePaths).toContain("docProps/app.xml");
    expect(opaquePaths).toContain("docProps/core.xml");
    for (const part of snap.root.opaqueParts.values()) {
      expect(part.bytes.length).toBeGreaterThan(0);
      expect(part.hash).toMatch(/^[0-9a-f]{64}$/);
    }
  });

  it("loads SheetJS workbook with cell data accessible via the escape hatch", async () => {
    const buf = await loadFixture("01-single-sheet-numbers.xlsx");
    const snap = await parseXlsx(buf);
    const ws = snap.root.sheetjs.Sheets["Inventory"];
    expect(ws).toBeDefined();
    const data = (ws as unknown as { "!data"?: unknown[][] })["!data"];
    expect(data).toBeDefined();
    const rows = data as Array<Array<{ v?: unknown }>>;
    expect(rows[0][0].v).toBe("Item");
    expect(rows[1][1].v).toBe(10);
    expect(rows[2][2].v).toBe(12.99);
  });

  it("parser is idempotent at the model boundary (modulo nodeId)", async () => {
    const buf = await loadFixture("04-merged-and-formatted.xlsx");
    const a = await parseXlsx(buf, { idMinter: deterministicIdMinter("a") });
    const b = await parseXlsx(buf, { idMinter: deterministicIdMinter("a") });
    expect(a.root.sheets.map((s) => s.name)).toEqual(b.root.sheets.map((s) => s.name));
    expect(a.root.sheets.map((s) => s.id)).toEqual(b.root.sheets.map((s) => s.id));
    expect(a.partHashes).toEqual(b.partHashes);
  });

  it("raises a typed error on a non-zip input", async () => {
    const garbage = new TextEncoder().encode("not a zip");
    await expect(parseXlsx(garbage)).rejects.toBeInstanceOf(XlsxParseError);
  });

  it("raises a typed error when xl/workbook.xml is missing", async () => {
    const JSZip = (await import("jszip")).default;
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<?xml version='1.0'?><Types/>");
    const buf = await zip.generateAsync({ type: "uint8array" });
    await expect(parseXlsx(buf)).rejects.toMatchObject({
      name: "XlsxParseError",
      code: "missing-workbook-part",
    });
  });
});

describe("resolveTargetPath", () => {
  it("resolves relative targets against the owner part directory", () => {
    expect(resolveTargetPath("xl/workbook.xml", "worksheets/sheet1.xml")).toBe("xl/worksheets/sheet1.xml");
  });

  it("resolves '../' segments correctly", () => {
    expect(resolveTargetPath("xl/worksheets/sheet1.xml", "../sharedStrings.xml")).toBe(
      "xl/sharedStrings.xml"
    );
    expect(resolveTargetPath("xl/worksheets/sheet1.xml", "../../docProps/app.xml")).toBe("docProps/app.xml");
  });

  it("treats absolute targets as workspace-rooted", () => {
    expect(resolveTargetPath("xl/workbook.xml", "/xl/sharedStrings.xml")).toBe("xl/sharedStrings.xml");
  });
});
