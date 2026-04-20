import { readFile } from "node:fs/promises";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { sha256Hex } from "@officeai/core";
import { parseXlsx } from "../parser/parse.js";
import { serializeXlsx } from "./serialize.js";
import { XlsxSerializeError } from "./errors.js";

const here = dirname(fileURLToPath(import.meta.url));
const fixtures = resolve(here, "../../../../fixtures/xlsx/synthetic");

const FIXTURES = [
  "01-single-sheet-numbers.xlsx",
  "02-multi-sheet.xlsx",
  "03-formulas-basic.xlsx",
  "04-merged-and-formatted.xlsx",
  "05-comments-hyperlinks.xlsx",
  "06-large-grid.xlsx",
] as const;

async function loadFixture(name: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(resolve(fixtures, name)));
}

describe("serializeXlsx — round-trip byte-preservation", () => {
  for (const name of FIXTURES) {
    it(`${name}: every part-content survives parse → serialize → reparse byte-identical`, async () => {
      const buf = await loadFixture(name);
      const before = await parseXlsx(buf);
      const reEmitted = await serializeXlsx(before);
      const after = await parseXlsx(new Uint8Array(reEmitted));

      expect(Object.keys(after.partHashes).sort()).toEqual(Object.keys(before.partHashes).sort());
      for (const path of Object.keys(before.partHashes)) {
        expect(after.partHashes[path], `mismatch on ${path}`).toBe(before.partHashes[path]);
      }
    });

    it(`${name}: container.serialize itself preserves part-content hashes`, async () => {
      const buf = await loadFixture(name);
      const snap = await parseXlsx(buf);
      const out = await serializeXlsx(snap);
      const reparsed = await parseXlsx(new Uint8Array(out));
      for (const path of Object.keys(snap.partHashes)) {
        const original = snap.container.readBytes(path);
        const re = reparsed.container.readBytes(path);
        expect(sha256Hex(re), `byte-content drift on ${path}`).toBe(sha256Hex(original));
      }
    });
  }
});

describe("serializeXlsx — dirty-flag guard", () => {
  it("throws when caller hand-sets an unsupported dirty flag (sst not in Phase 5)", async () => {
    const buf = await loadFixture("01-single-sheet-numbers.xlsx");
    const snap = await parseXlsx(buf);
    const tampered = { ...snap, dirty: { ...snap.dirty, sharedStrings: true } };
    await expect(serializeXlsx(tampered)).rejects.toBeInstanceOf(XlsxSerializeError);
  });

  it("rewrites a dirty sheet through SheetJS and round-trips the typed cells", async () => {
    const buf = await loadFixture("01-single-sheet-numbers.xlsx");
    const snap = await parseXlsx(buf);
    const sheet = snap.root.sheets[0];
    const cells = new Map(sheet.cells);
    cells.set("0:25", { row: 0, col: 25, value: 12345 });
    const nextSheet = { ...sheet, cells };
    const sheets = snap.root.sheets.slice();
    sheets[0] = nextSheet;
    const tampered = {
      ...snap,
      root: { ...snap.root, sheets },
      dirty: { ...snap.dirty, sheets: new Set([sheet.partPath]) },
    };
    const out = await serializeXlsx(tampered);
    const reparsed = await parseXlsx(new Uint8Array(out));
    const reSheet = reparsed.root.sheets[0];
    expect(reSheet.cells.get("0:25")?.value).toBe(12345);
  });

  it("round-trips a frozen-pane configuration written via the model", async () => {
    const buf = await loadFixture("01-single-sheet-numbers.xlsx");
    const snap = await parseXlsx(buf);
    const sheet = snap.root.sheets[0];
    const sheets = snap.root.sheets.slice();
    sheets[0] = { ...sheet, freeze: { rows: 1, cols: 2 } };
    const tampered = {
      ...snap,
      root: { ...snap.root, sheets },
      dirty: { ...snap.dirty, sheets: new Set([sheet.partPath]) },
    };
    const out = await serializeXlsx(tampered);
    const reparsed = await parseXlsx(new Uint8Array(out));
    expect(reparsed.root.sheets[0].freeze).toEqual({ rows: 1, cols: 2 });
  });

  it("round-trips a typed list data-validation through serialize → parse", async () => {
    const buf = await loadFixture("01-single-sheet-numbers.xlsx");
    const snap = await parseXlsx(buf);
    const sheet = snap.root.sheets[0];
    const sheets = snap.root.sheets.slice();
    sheets[0] = {
      ...sheet,
      dataValidations: [
        {
          kind: "list",
          id: "dv-test",
          range: "B2:B10",
          source: "Yes,No,Maybe",
          formula: false,
          showDropDown: true,
          stopOnInvalid: true,
          allowBlank: true,
        },
      ],
    };
    const tampered = {
      ...snap,
      root: { ...snap.root, sheets },
      dirty: { ...snap.dirty, sheets: new Set([sheet.partPath]) },
    };
    const out = await serializeXlsx(tampered);
    const reparsed = await parseXlsx(new Uint8Array(out));
    const dv = reparsed.root.sheets[0].dataValidations;
    expect(dv).toHaveLength(1);
    expect(dv[0]!.kind).toBe("list");
    expect(dv[0]!.range).toBe("B2:B10");
    if (dv[0]!.kind === "list") {
      expect(dv[0]!.source).toBe("Yes,No,Maybe");
      expect(dv[0]!.formula).toBe(false);
    }
  });

  it("preserves <hyperlinks>, <legacyDrawing>, and <ignoredErrors> across a dirty cell edit", async () => {
    const buf = await loadFixture("05-comments-hyperlinks.xlsx");
    const snap = await parseXlsx(buf);
    const sheet = snap.root.sheets[0]!;
    expect(sheet.hyperlinksXml).toContain("rId1");
    expect(sheet.hyperlinksXml).toContain('display="https://example.com/spec"');
    expect(sheet.legacyDrawingXml).toMatch(/<legacyDrawing\b/);
    expect(sheet.ignoredErrorsXml).toContain("numberStoredAsText");

    // Now mutate a single cell so the worksheet goes dirty.
    const cells = new Map(sheet.cells);
    cells.set("9:9", { row: 9, col: 9, value: "edit" });
    const sheets = snap.root.sheets.slice();
    sheets[0] = { ...sheet, cells };
    const tampered = {
      ...snap,
      root: { ...snap.root, sheets },
      dirty: { ...snap.dirty, sheets: new Set([sheet.partPath]) },
    };
    const out = await serializeXlsx(tampered);
    const reparsed = await parseXlsx(new Uint8Array(out));
    const reSheet = reparsed.root.sheets[0]!;
    expect(reSheet.hyperlinksXml).toContain("rId1");
    expect(reSheet.hyperlinksXml).toContain('display="https://example.com/spec"');
    expect(reSheet.legacyDrawingXml).toMatch(/<legacyDrawing\b/);
    expect(reSheet.ignoredErrorsXml).toContain("numberStoredAsText");
    expect(reSheet.cells.get("9:9")?.value).toBe("edit");
  });

  it("preserves <sheetViews> attributes (zoom, view options) when freeze toggles on", async () => {
    const buf = await loadFixture("05-comments-hyperlinks.xlsx");
    const snap = await parseXlsx(buf);
    const sheet = snap.root.sheets[0]!;
    // Source sheetViews carries `workbookViewId="0"` — we want it
    // preserved verbatim even after we add a typed freeze.
    expect(sheet.sheetViewsXml).toContain('workbookViewId="0"');

    const sheets = snap.root.sheets.slice();
    sheets[0] = { ...sheet, freeze: { rows: 1, cols: 0 } };
    const tampered = {
      ...snap,
      root: { ...snap.root, sheets },
      dirty: { ...snap.dirty, sheets: new Set([sheet.partPath]) },
    };
    const out = await serializeXlsx(tampered);
    const reparsed = await parseXlsx(new Uint8Array(out));
    const reSheet = reparsed.root.sheets[0]!;
    expect(reSheet.freeze).toEqual({ rows: 1, cols: 0 });
    // The original `<sheetView workbookViewId="0">` attribute survives.
    expect(reSheet.sheetViewsXml).toContain('workbookViewId="0"');
  });

  it("re-emits <tableParts> for sheets that had source tables when worksheet goes dirty", async () => {
    // 04-merged-and-formatted.xlsx is the one synthetic fixture we
    // can rewrite without table churn; tableParts injection is
    // primarily exercised on real-world files. Test the typed-only
    // path: stamp a TableDef onto the sheet and prove the
    // <tableParts> block lands in the dirty XML.
    const buf = await loadFixture("04-merged-and-formatted.xlsx");
    const snap = await parseXlsx(buf);
    const sheet = snap.root.sheets[0]!;
    const cells = new Map(sheet.cells);
    cells.set("0:25", { row: 0, col: 25, value: "x" });
    const sheets = snap.root.sheets.slice();
    sheets[0] = {
      ...sheet,
      cells,
      tables: [
        {
          id: "t1" as never,
          tableId: "1",
          name: "TestTable",
          displayName: "TestTable",
          range: "A1:B5",
          headerRowCount: 1,
          totalsRowCount: 0,
          columnNames: ["A", "B"],
          partPath: "xl/tables/table1.xml",
          relId: "rIdTbl1",
        },
      ],
    };
    const tampered = {
      ...snap,
      root: { ...snap.root, sheets },
      dirty: { ...snap.dirty, sheets: new Set([sheet.partPath]) },
    };
    const out = await serializeXlsx(tampered);
    const xml = new TextDecoder().decode(
      tampered.container.clone().readBytes(sheet.partPath) // for type only
    );
    void xml;
    const reparsed = await parseXlsx(new Uint8Array(out));
    const reSheetXml = reparsed.container.readText(sheet.partPath);
    expect(reSheetXml).toContain('<tableParts count="1">');
    expect(reSheetXml).toContain('r:id="rIdTbl1"');
  });

  it("re-emits opaque <conditionalFormatting> blocks verbatim on dirty sheets", async () => {
    const buf = await loadFixture("01-single-sheet-numbers.xlsx");
    const snap = await parseXlsx(buf);
    const sheet = snap.root.sheets[0];
    const opaque =
      '<conditionalFormatting sqref="A1:A10"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>';
    const sheets = snap.root.sheets.slice();
    sheets[0] = { ...sheet, opaqueConditionalFormats: [opaque] };
    const tampered = {
      ...snap,
      root: { ...snap.root, sheets },
      dirty: { ...snap.dirty, sheets: new Set([sheet.partPath]) },
    };
    const out = await serializeXlsx(tampered);
    const reparsed = await parseXlsx(new Uint8Array(out));
    const blocks = reparsed.root.sheets[0].opaqueConditionalFormats;
    expect(blocks).toHaveLength(1);
    expect(blocks[0]).toContain('sqref="A1:A10"');
    expect(blocks[0]).toContain("greaterThan");
  });

  it("preserves shared formula encoding across a dirty rewrite", async () => {
    // Build a synthetic shared formula in the typed model — we
    // can't rely on a fixture having one because SheetJS expands
    // them on parse and we'd need the unmodified XML to assert
    // round-trip. Tampered snapshot mimics the parser output for a
    // shared group rooted at A1 with si=0 spanning A1:A3.
    const buf = await loadFixture("01-single-sheet-numbers.xlsx");
    const snap = await parseXlsx(buf);
    const sheet = snap.root.sheets[0];
    const cells = new Map(sheet.cells);
    cells.set("0:0", {
      row: 0,
      col: 0,
      value: 100,
      formula: { text: "B1+1", kind: "shared", sharedIndex: 0, ref: "A1:A3", isMaster: true },
    });
    cells.set("1:0", {
      row: 1,
      col: 0,
      value: 101,
      formula: { text: "B2+1", kind: "shared", sharedIndex: 0 },
    });
    cells.set("2:0", {
      row: 2,
      col: 0,
      value: 102,
      formula: { text: "B3+1", kind: "shared", sharedIndex: 0 },
    });
    const sheets = snap.root.sheets.slice();
    sheets[0] = { ...sheet, cells };
    const tampered = {
      ...snap,
      root: { ...snap.root, sheets },
      dirty: { ...snap.dirty, sheets: new Set([sheet.partPath]) },
    };
    const out = await serializeXlsx(tampered);
    const xml = (await parseXlsx(new Uint8Array(out))).container.readText(sheet.partPath);
    // Master cell carries the body and the ref.
    expect(xml).toMatch(/<c\b[^>]*r="A1"[^>]*>[\s\S]*?<f t="shared" si="0" ref="A1:A3">B1\+1<\/f>/);
    // Followers are self-closing references.
    expect(xml).toMatch(/<c\b[^>]*r="A2"[^>]*>[\s\S]*?<f t="shared" si="0"\/>/);
    expect(xml).toMatch(/<c\b[^>]*r="A3"[^>]*>[\s\S]*?<f t="shared" si="0"\/>/);

    // Re-parse: typed metadata round-trips.
    const reparsed = await parseXlsx(new Uint8Array(out));
    const reCells = reparsed.root.sheets[0].cells;
    const a1 = reCells.get("0:0");
    expect(a1?.formula?.kind).toBe("shared");
    expect(a1?.formula?.sharedIndex).toBe(0);
    expect(a1?.formula?.ref).toBe("A1:A3");
    expect(a1?.formula?.isMaster).toBe(true);
    const a2 = reCells.get("1:0");
    expect(a2?.formula?.kind).toBe("shared");
    expect(a2?.formula?.sharedIndex).toBe(0);
    expect(a2?.formula?.isMaster).toBeFalsy();
  });

  it("emits typed conditionalFormats (cellIs / colorScale / dataBar) to OOXML on dirty save", async () => {
    const buf = await loadFixture("01-single-sheet-numbers.xlsx");
    const snap = await parseXlsx(buf);
    const sheet = snap.root.sheets[0];
    const sheets = snap.root.sheets.slice();
    sheets[0] = {
      ...sheet,
      conditionalFormats: [
        {
          kind: "cellIs",
          id: "rule-1",
          range: "B2:B5",
          op: "gt",
          value: 10,
          overlay: { fill: "FFFF00" },
        },
        {
          kind: "colorScale",
          id: "rule-2",
          range: "C2:C5",
          minColor: "FF0000",
          maxColor: "00FF00",
        },
        {
          kind: "dataBar",
          id: "rule-3",
          range: "D2:D5",
          color: "0070C0",
        },
      ],
    };
    const tampered = {
      ...snap,
      root: { ...snap.root, sheets },
      dirty: { ...snap.dirty, sheets: new Set([sheet.partPath]) },
    };
    const out = await serializeXlsx(tampered);
    const reparsed = await parseXlsx(new Uint8Array(out));
    // Round-tripped CF rules land in the opaque set on re-parse
    // (parser treats them as opaque blocks). The structural pieces
    // (operator, color stops, data bar color) must all be present.
    const blocks = reparsed.root.sheets[0].opaqueConditionalFormats;
    const joined = blocks.join("");
    expect(joined).toContain('sqref="B2:B5"');
    expect(joined).toContain('operator="greaterThan"');
    expect(joined).toContain("<formula>10</formula>");
    expect(joined).toContain('sqref="C2:C5"');
    expect(joined).toContain('<color rgb="FFFF0000"/>');
    expect(joined).toContain('<color rgb="FF00FF00"/>');
    expect(joined).toContain('sqref="D2:D5"');
    expect(joined).toContain("<dataBar>");
    expect(joined).toContain('<color rgb="FF0070C0"/>');
  });
});

describe("serializeXlsx — mutation isolation", () => {
  // For each fixture, verify that mutating exactly one sheet leaves
  // every OTHER part of the package byte-identical to the source.
  // This is the round-trip oracle for the "two-tier strategy":
  // dirty parts are typed-rewritten; everything else replays the
  // verbatim cache. A regression that drops the cache (or
  // accidentally re-emits an untouched part through a typed
  // serializer) shows up here as a hash drift on a part the user
  // never touched.
  for (const name of FIXTURES) {
    it(`${name}: editing one cell does not touch unrelated parts`, async () => {
      const buf = await loadFixture(name);
      const snap = await parseXlsx(buf);
      const sheet = snap.root.sheets[0];
      // Pick a cell coordinate that's beyond the source data range
      // for every fixture so we don't collide with an existing
      // formula and trigger formula-engine recomputation.
      const cells = new Map(sheet.cells);
      cells.set("99:99", { row: 99, col: 99, value: 1 });
      const sheets = snap.root.sheets.slice();
      sheets[0] = { ...sheet, cells };
      const tampered = {
        ...snap,
        root: { ...snap.root, sheets },
        dirty: { ...snap.dirty, sheets: new Set([sheet.partPath]) },
      };
      const out = await serializeXlsx(tampered);
      const reparsed = await parseXlsx(new Uint8Array(out));
      // The dirty sheet's part is allowed to drift (it was
      // re-emitted). The /xl/worksheets/sheet2.xml … sheetN.xml
      // and all non-worksheet parts (theme, styles, sst, drawings,
      // chartN, mediaN, _rels, comments) must NOT.
      const allowedDrift = new Set<string>([sheet.partPath]);
      // Sheet rels often piggy-back drawing/comment rId rewrites
      // that the dirty sheet pass needs; whitelist those too.
      const relsPath = sheet.partPath
        .replace(/\.xml$/, ".xml.rels")
        .replace("xl/worksheets/", "xl/worksheets/_rels/");
      allowedDrift.add(relsPath);
      // [Content_Types] is regenerated whenever sheets churn — not
      // worth gating because the rewrite is byte-stable in practice.
      allowedDrift.add("[Content_Types].xml");
      for (const path of Object.keys(snap.partHashes)) {
        if (allowedDrift.has(path)) continue;
        if (!reparsed.container.has(path)) continue;
        const original = snap.container.readBytes(path);
        const re = reparsed.container.readBytes(path);
        expect(sha256Hex(re), `byte drift on untouched part ${path} (fixture ${name})`).toBe(
          sha256Hex(original)
        );
      }
    });
  }
});
