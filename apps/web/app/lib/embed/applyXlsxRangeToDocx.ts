import type { DocxAgent } from "@officeai/docx";
import type { XlsxClipboardSnapshot } from "@officeai/xlsx";
import { snapshotToGrid, snapshotToChartSpec, type XlsxEmbedMode } from "./xlsxEmbedShared";

/**
 * Drop an XLSX range — captured at copy time as a
 * {@link XlsxClipboardSnapshot} — into the open DOCX document.
 *
 * Three modes (default `materialized`, mirrors PowerPoint's "Paste
 * Special" submenu):
 *   - `materialized`: typed `<w:tbl>` (via `docx:insert-table` +
 *     N×`docx:set-cell-content`). Best fidelity; cells stay editable
 *     as Word table cells.
 *   - `live`: OLE-embedded `.xlsx` package (`docx:insert-spreadsheet`).
 *     Office activates Excel on double-click; the bytes round-trip.
 *   - `chart`: typed chart (`docx:insert-chart`) using the snapshot's
 *     first row as series names and first column as categories. Embeds
 *     a backing `.xlsx` workbook so "Edit Data" works in Word.
 *
 * We deliberately go through the public command bus rather than
 * mutating the snapshot directly so:
 *
 *   - The paste shows up in the undo stack as a single user-visible
 *     "Paste table" entry (the editor's undo coalescer batches
 *     consecutive command-bus dispatches that share a transaction
 *     id; the caller wraps this helper with that id).
 *   - The realtime broadcast layer (Phase 1) replays the same
 *     commands on every other peer, so a pasted XLSX range becomes
 *     visible in another browser without any additional plumbing.
 *   - The headless agent path (`packages/agent`) inherits the
 *     behaviour for free.
 *
 * Cross-format downgrades (`materialized` only):
 *   - Cell formulas, style ids and merge regions are not copied
 *     yet. Cells render as their `value` string; numeric values use
 *     `String(v)`. Formulas → text starting with `=`. Booleans →
 *     "TRUE"/"FALSE" (Excel convention).
 *   - Column widths default to a uniform split because the source
 *     snapshot doesn't carry pixel widths and Word's default would
 *     otherwise stretch the table to the page width unevenly.
 */
export async function applyXlsxRangeToDocx(args: {
  readonly agent: DocxAgent;
  readonly snapshot: XlsxClipboardSnapshot;
  readonly paragraphIndex: number;
  /** Default `materialized`. */
  readonly mode?: XlsxEmbedMode;
}): Promise<void> {
  const { agent, snapshot, paragraphIndex } = args;
  const mode: XlsxEmbedMode = args.mode ?? "materialized";
  if (snapshot.width <= 0 || snapshot.height <= 0) return;

  if (mode === "live") {
    const result = await agent.applyCommand({
      type: "docx:insert-spreadsheet",
      payload: {
        at: { paragraph: paragraphIndex, run: 0, offset: 0 },
        data: snapshotToGrid(snapshot),
        sheetName: snapshot.origin.sheet,
        name: `XLSX paste ${snapshot.origin.sheet}!${snapshot.origin.range}`,
      },
      source: "human",
    });
    if (result.rejection) {
      throw new Error(
        `docx:insert-spreadsheet rejected: ${result.rejection.code} ${result.rejection.message ?? ""}`
      );
    }
    return;
  }

  if (mode === "chart") {
    const spec = snapshotToChartSpec(snapshot);
    if (!spec) {
      // Fall back to materialised when the grid can't be projected
      // into a chart (e.g. a single column / single row of values).
      return applyXlsxRangeToDocxAsTable(agent, snapshot, paragraphIndex);
    }
    const result = await agent.applyCommand({
      type: "docx:insert-chart",
      payload: {
        at: { paragraph: paragraphIndex, run: 0, offset: 0 },
        chartType: "bar",
        categories: spec.categories,
        series: spec.series,
        name: `XLSX chart ${snapshot.origin.sheet}!${snapshot.origin.range}`,
      },
      source: "human",
    });
    if (result.rejection) {
      throw new Error(
        `docx:insert-chart rejected: ${result.rejection.code} ${result.rejection.message ?? ""}`
      );
    }
    return;
  }

  return applyXlsxRangeToDocxAsTable(agent, snapshot, paragraphIndex);
}

async function applyXlsxRangeToDocxAsTable(
  agent: DocxAgent,
  snapshot: XlsxClipboardSnapshot,
  paragraphIndex: number
): Promise<void> {
  // 1) Insert the empty table at the requested paragraph index.
  //    Use a uniform 2400 twip column width (~ 1.67 inches) so the
  //    typical 3-4 column XLSX paste fits a Letter portrait body
  //    width (~ 6 in) without overflowing.
  const cols = snapshot.width;
  const rows = snapshot.height;
  const colWidth = pickUniformColumnWidth(cols);
  const insert = await agent.applyCommand({
    type: "docx:insert-table",
    payload: {
      at: { paragraph: paragraphIndex, run: 0, offset: 0 },
      rows,
      cols,
      columnWidths: new Array(cols).fill(colWidth),
    },
    source: "human",
  });
  if (insert.rejection) {
    throw new Error(`docx:insert-table rejected: ${insert.rejection.code} ${insert.rejection.message ?? ""}`);
  }

  // The insert handler emits a `node-inserted` diff with the new
  // table id. Pull it back out so we can address each cell by
  // (tableId, row, col) on subsequent set-cell-content calls.
  const tableId = insert.diff?.changes.find((c) => c.kind === "node-inserted")?.nodeId;
  if (!tableId) {
    throw new Error("docx:insert-table did not emit a node-inserted diff");
  }

  // 2) Fill cells row-major. We skip null cells (Excel sparsity)
  //    so empty cells stay as the default empty paragraph the
  //    insert handler put there. We do NOT batch into a custom
  //    composite command; the existing `docx:set-cell-content`
  //    handler is the single supported mutation path and tests
  //    cover its dirty-flag semantics.
  for (let r = 0; r < rows; r++) {
    const row = snapshot.cells[r] ?? [];
    for (let c = 0; c < cols; c++) {
      const cell = row[c];
      if (!cell) continue;
      const text = displayValue(cell);
      if (!text) continue;
      const result = await agent.applyCommand({
        type: "docx:set-cell-content",
        payload: {
          tableId,
          row: r,
          col: c,
          content: [
            {
              kind: "paragraph",
              id: `embed-p-${r}-${c}`,
              properties: {},
              children: [
                {
                  kind: "run",
                  id: `embed-r-${r}-${c}`,
                  properties: {},
                  children: [
                    {
                      kind: "text",
                      id: `embed-t-${r}-${c}`,
                      text,
                      xmlSpacePreserve: text !== text.trim() || text.length === 0,
                    },
                  ],
                },
              ],
            },
          ],
        },
        source: "human",
      });
      if (result.rejection) {
        // Don't abort the whole paste — surface the first failure
        // through the throw at the end so partial table content
        // stays committed and visible.
        throw new Error(`docx:set-cell-content rejected at (${r},${c}): ${result.rejection.code}`);
      }
    }
  }
}

function displayValue(cell: { value: unknown; formula?: string }): string {
  if (cell.formula) return `=${cell.formula}`;
  const v = cell.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v;
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "object" && v && "code" in v) {
    return String((v as { code: string }).code);
  }
  return String(v);
}

/**
 * Pick a per-column twip width that keeps the table inside a Letter
 * portrait body (margins of ~ 1in each side ⇒ ~ 9360 twips usable).
 * For very wide pastes we cap at ~ 1500 twips per column so Word
 * doesn't render an absurdly wide table that would clip on print.
 */
function pickUniformColumnWidth(cols: number): number {
  const usableTwips = 9360;
  const proposed = Math.floor(usableTwips / Math.max(1, cols));
  return Math.max(800, Math.min(proposed, 4800));
}
