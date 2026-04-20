import type { XlsxAgent } from "@officeai/xlsx";
import type { XlsxClipboardSnapshot } from "@officeai/xlsx";

/**
 * D5 — drop a DOCX-table envelope into the open XLSX workbook as a
 * real range. Each row of `cells` becomes a sheet row anchored at
 * `target` (an A1 single-cell ref); the values keep XLSX paste's
 * type-inference conventions:
 *
 *   - Empty strings become sparse `null` cells (so the paste doesn't
 *     overwrite the destination's existing values with blanks).
 *   - Strings that match `^-?\d+(?:\.\d+)?$` become `number` cells —
 *     mirrors the in-app HTML / TSV importer in
 *     `packages/xlsx/src/clipboard/external.ts`.
 *   - `"TRUE"` / `"FALSE"` (case-insensitive, Excel's serialised
 *     boolean form) become `boolean` cells.
 *   - Everything else stays a string.
 *
 * We dispatch through the existing `xlsx:paste-range` command so the
 * paste participates in the same undo unit, realtime broadcast, and
 * marching-ants logic as a regular range paste — no new command
 * surface needed.
 */
export async function applyDocxTableToXlsx(args: {
  readonly agent: XlsxAgent;
  readonly sheet: string;
  readonly target: string;
  readonly cells: ReadonlyArray<ReadonlyArray<string>>;
  readonly originLabel?: string;
}): Promise<void> {
  const { agent, sheet, target, cells } = args;
  if (cells.length === 0) return;
  const width = cells.reduce((m, r) => Math.max(m, r.length), 0);
  if (width === 0) return;
  const snapshot = buildSnapshotFromMatrix({
    sheet,
    target,
    width,
    rows: cells,
  });
  const result = await agent.applyCommand({
    type: "xlsx:paste-range",
    payload: {
      sheet,
      target,
      source: snapshot,
      mode: "all",
      transpose: false,
    },
    source: "human",
  });
  if (result.rejection) {
    throw new Error(
      `xlsx:paste-range rejected for docx-table embed: ${result.rejection.code}` +
        (result.rejection.message ? ` ${result.rejection.message}` : "")
    );
  }
}

function buildSnapshotFromMatrix(args: {
  sheet: string;
  target: string;
  width: number;
  rows: ReadonlyArray<ReadonlyArray<string>>;
}): XlsxClipboardSnapshot {
  const { sheet, target, width, rows } = args;
  const height = rows.length;
  const cells = rows.map((row) =>
    Array.from({ length: width }, (_, c) => coerceCell(row[c] ?? ""))
  );
  return {
    origin: { sheet, range: target },
    width,
    height,
    cells,
    merges: [],
  };
}

function coerceCell(raw: string): { value: string | number | boolean } | null {
  const trimmed = raw.trim();
  if (trimmed === "") return null;
  if (/^-?\d+(?:\.\d+)?$/.test(trimmed)) {
    const n = Number(trimmed);
    if (Number.isFinite(n)) return { value: n };
  }
  const lower = trimmed.toLowerCase();
  if (lower === "true") return { value: true };
  if (lower === "false") return { value: false };
  return { value: raw };
}
