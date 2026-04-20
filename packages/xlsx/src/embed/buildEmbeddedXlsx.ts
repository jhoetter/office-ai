import JSZip from "jszip";

/**
 * The 2D data grid we embed inside Word/PowerPoint charts. The first row
 * is treated as series-name headers (cell A1 is empty / the category
 * column header), the first column as category labels, and the remaining
 * cells as numeric series values.
 *
 * Example for a chart with two series ("Q1", "Q2") and three categories
 * ("Apples", "Oranges", "Pears"):
 *
 *   [
 *     ["",        "Q1", "Q2"],
 *     ["Apples",  10,   20  ],
 *     ["Oranges", 15,   25  ],
 *     ["Pears",   12,   18  ],
 *   ]
 */
export type EmbeddedGridCell = string | number | null | undefined;
export type EmbeddedGrid = ReadonlyArray<ReadonlyArray<EmbeddedGridCell>>;

export interface BuildEmbeddedXlsxOptions {
  /** Worksheet name. Defaults to `"Sheet1"`. */
  readonly sheetName?: string;
}

export interface BuildEmbeddedXlsxResult {
  /** Bytes of a fully-valid `.xlsx` package. */
  readonly bytes: Uint8Array;
  /** Worksheet name actually used. */
  readonly sheetName: string;
  /** Number of data rows including the header. */
  readonly rowCount: number;
  /** Number of columns including the category column. */
  readonly columnCount: number;
}

/**
 * Build a minimal but spec-valid `.xlsx` package whose single worksheet
 * is populated from the given grid. Designed for embedding inside a
 * DrawingML chart (`c:externalData` / `…/embeddings/Microsoft_Excel_Worksheet.xlsx`)
 * or as the payload of an OLE-Excel object inside a Word/PowerPoint
 * package.
 *
 * The output passes Excel/PowerPoint's "Edit Data" round-trip cleanly:
 *   - `[Content_Types].xml` overrides for workbook + worksheet
 *   - `_rels/.rels` -> `xl/workbook.xml`
 *   - `xl/_rels/workbook.xml.rels` -> `xl/worksheets/sheet1.xml`
 *   - `xl/workbook.xml` declaring one sheet with the requested name
 *   - `xl/worksheets/sheet1.xml` with `<sheetData>` rows, inline strings
 *     for non-numeric cells, plain `<v>` for numbers
 *
 * No shared-strings table, no styles, no theme — Office happily reads
 * inline strings (`<is><t>…</t></is>`) without them.
 */
export async function buildEmbeddedXlsx(
  grid: EmbeddedGrid,
  opts: BuildEmbeddedXlsxOptions = {}
): Promise<BuildEmbeddedXlsxResult> {
  const sheetName = opts.sheetName ?? "Sheet1";
  validateSheetName(sheetName);

  const rowCount = grid.length;
  const columnCount = grid.reduce((max, row) => Math.max(max, row.length), 0);

  const sheetXml = buildSheetXml(grid);
  const workbookXml = buildWorkbookXml(sheetName);

  const z = new JSZip();
  z.file("[Content_Types].xml", CONTENT_TYPES_XML);
  z.file("_rels/.rels", PACKAGE_RELS_XML);
  z.file("xl/workbook.xml", workbookXml);
  z.file("xl/_rels/workbook.xml.rels", WORKBOOK_RELS_XML);
  z.file("xl/worksheets/sheet1.xml", sheetXml);

  const bytes = (await z.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  })) as Uint8Array;

  return { bytes, sheetName, rowCount, columnCount };
}

/**
 * Build a chart-shaped grid from a category list and an arbitrary
 * number of data series. The result is laid out so the first row holds
 * series names (with `A1` left empty), the first column holds the
 * category labels, and the rest of the cells are numeric series values.
 *
 * Returned together with the cell-reference strings that the chart's
 * `<c:f>` elements should point at — guaranteeing the embedded
 * workbook's layout exactly matches what the chart claims to read.
 */
export interface ChartGridSeries {
  readonly name?: string;
  readonly values: ReadonlyArray<number>;
}

export interface ChartGridResult {
  readonly grid: EmbeddedGrid;
  readonly sheetName: string;
  /** `Sheet1!$A$2:$A$N` — the category cells (first column, header row excluded). */
  readonly categoryRef: string;
  /** Per-series `Sheet1!$<col>$2:$<col>$N` value-range refs. */
  readonly valueRefs: ReadonlyArray<string>;
  /** Per-series `Sheet1!$<col>$1` series-name cell refs. */
  readonly nameRefs: ReadonlyArray<string>;
}

export function buildChartGrid(
  categories: ReadonlyArray<string>,
  series: ReadonlyArray<ChartGridSeries>,
  opts: { sheetName?: string } = {}
): ChartGridResult {
  const sheetName = opts.sheetName ?? "Sheet1";
  validateSheetName(sheetName);

  const header: EmbeddedGridCell[] = [""];
  for (let i = 0; i < series.length; i++) {
    header.push(series[i]!.name ?? `Series ${i + 1}`);
  }
  const rows: EmbeddedGridCell[][] = [header];
  for (let r = 0; r < categories.length; r++) {
    const row: EmbeddedGridCell[] = [categories[r] ?? ""];
    for (let s = 0; s < series.length; s++) {
      const v = series[s]!.values[r];
      row.push(typeof v === "number" && Number.isFinite(v) ? v : 0);
    }
    rows.push(row);
  }

  const lastDataRow = categories.length + 1;
  const categoryRef = `${quoteSheetName(sheetName)}!$A$2:$A$${lastDataRow}`;
  const valueRefs: string[] = [];
  const nameRefs: string[] = [];
  for (let s = 0; s < series.length; s++) {
    const col = colToLetter(s + 2);
    valueRefs.push(`${quoteSheetName(sheetName)}!$${col}$2:$${col}$${lastDataRow}`);
    nameRefs.push(`${quoteSheetName(sheetName)}!$${col}$1`);
  }

  return { grid: rows, sheetName, categoryRef, valueRefs, nameRefs };
}

// ─── XML builders ─────────────────────────────────────────────────────────

function buildSheetXml(grid: EmbeddedGrid): string {
  const rowsXml: string[] = [];
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r] ?? [];
    const cellsXml: string[] = [];
    for (let c = 0; c < row.length; c++) {
      const v = row[c];
      if (v === null || v === undefined || v === "") continue;
      const ref = `${colToLetter(c + 1)}${r + 1}`;
      if (typeof v === "number" && Number.isFinite(v)) {
        cellsXml.push(`<c r="${ref}"><v>${v}</v></c>`);
      } else {
        cellsXml.push(
          `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(String(v))}</t></is></c>`
        );
      }
    }
    if (cellsXml.length === 0) continue;
    rowsXml.push(`<row r="${r + 1}">${cellsXml.join("")}</row>`);
  }
  const dimension = grid.length === 0 ? "A1" : `A1:${colToLetter(Math.max(1, maxCols(grid)))}${grid.length}`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="${dimension}"/>
<sheetData>${rowsXml.join("")}</sheetData>
</worksheet>`;
}

function buildWorkbookXml(sheetName: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="${escapeXml(sheetName)}" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>`;
}

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;

const PACKAGE_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

const WORKBOOK_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;

// ─── Utilities ────────────────────────────────────────────────────────────

function maxCols(grid: EmbeddedGrid): number {
  let m = 0;
  for (const row of grid) m = Math.max(m, row.length);
  return m;
}

function colToLetter(col: number): string {
  let s = "";
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s || "A";
}

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function quoteSheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

function validateSheetName(name: string): void {
  if (!name) throw new Error("sheet name is required");
  if (name.length > 31) throw new Error(`sheet name too long (max 31 chars): ${name}`);
  if (/[\\/?*[\]:]/.test(name)) {
    throw new Error(`sheet name contains forbidden character: ${name}`);
  }
}
