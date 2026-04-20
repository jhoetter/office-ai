import type { PptxAgent } from "@officeai/pptx";
import type { CellValue, XlsxClipboardCell, XlsxClipboardSnapshot } from "@officeai/xlsx";

/**
 * Drop an XLSX range into PPTX as a freshly-added table on the
 * current slide.
 *
 * We dispatch a single `pptx:insert-table` so the paste materialises
 * as a real `<p:graphicFrame>/<a:tbl>` graphic frame — that means
 * PowerPoint and LibreOffice both render it as a table (cells stay
 * editable, formatting can be tweaked without losing row/column
 * structure) instead of the previous TSV-text-box compromise.
 *
 * The handler picks a sensibly centred default frame when we omit
 * `xEmu/yEmu/widthEmu/heightEmu`, so we don't have to query the deck
 * for slide dimensions here. Per-cell styles are intentionally NOT
 * forwarded yet: the PPTX command only accepts cell text strings,
 * and expanding its surface area is a separate piece of work.
 */
export async function applyXlsxRangeToPptx(args: {
  readonly agent: PptxAgent;
  readonly snapshot: XlsxClipboardSnapshot;
  readonly slideIndex: number;
}): Promise<void> {
  const { agent, snapshot, slideIndex } = args;
  if (snapshot.width <= 0 || snapshot.height <= 0) return;

  const cells = buildCellMatrix(snapshot);

  const result = await agent.applyCommand({
    type: "pptx:insert-table",
    payload: {
      slideIndex,
      rows: snapshot.height,
      cols: snapshot.width,
      cells,
      name: `XLSX paste ${snapshot.origin.sheet}!${snapshot.origin.range}`,
    },
    source: "human",
  });
  if (result.rejection) {
    throw new Error(
      `pptx:insert-table rejected: ${result.rejection.code} ${result.rejection.message ?? ""}`
    );
  }
}

/**
 * Project the snapshot's row-major cells into a `string[][]` matrix
 * sized exactly `height × width`. Empty / null positions become
 * `""` (NOT the literal `"undefined"` — the table command would
 * happily render the latter as a paragraph).
 */
function buildCellMatrix(snap: XlsxClipboardSnapshot): string[][] {
  const out: string[][] = [];
  for (let r = 0; r < snap.height; r++) {
    const row = snap.cells[r] ?? [];
    const cols: string[] = new Array(snap.width);
    for (let c = 0; c < snap.width; c++) {
      const cell = row[c];
      cols[c] = cell ? displayCell(cell) : "";
    }
    out.push(cols);
  }
  return out;
}

/**
 * Render one cell to display text.
 *
 * - Formula cells prefer the cached result (`cell.value`) when one is
 *   available, falling back to the raw formula text (with a leading
 *   `=`) so an unevaluated formula still leaves a visible breadcrumb
 *   instead of a blank cell.
 * - Non-formula cells use the same value-formatting rules the TSV
 *   path uses (`snapshotToTsv`): numbers via `String(n)`, booleans as
 *   `TRUE`/`FALSE`, errors as their `#code`.
 */
function displayCell(cell: XlsxClipboardCell): string {
  if (cell.formula) {
    if (cell.value !== null && cell.value !== undefined) {
      return formatValue(cell.value);
    }
    return `=${cell.formula}`;
  }
  return formatValue(cell.value);
}

function formatValue(value: CellValue): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "object" && "kind" in value && value.kind === "error") {
    return value.code;
  }
  return String(value);
}
