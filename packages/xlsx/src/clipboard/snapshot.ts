import { cellKey, formatRange, parseA1, parseRange } from "../model/refs.js";
import { flattenCellXf, type EffectiveStyle } from "../model/style-mutate.js";
import type { StyleTable } from "../model/style-table.js";
import type { CellValue, Sheet } from "../model/types.js";

/**
 * One cell in a `XlsxClipboardSnapshot`. Either a value-only cell, a
 * formula cell, or a styled empty cell. `null` rows are sparse —
 * the caller still indexes into the matrix with `rows[r][c]` but the
 * value is `null` for empty positions.
 */
export interface XlsxClipboardCell {
  readonly value: CellValue;
  /**
   * Raw formula text WITHOUT a leading `=`. When present, the
   * formula is the source of truth — pasting it relative-shifts the
   * refs against the destination top-left.
   */
  readonly formula?: string;
  /**
   * Workbook-scoped style index. Pasted only when the destination
   * uses the same style table; otherwise the paste handler resolves
   * an `EffectiveStyle` from the source workbook before re-interning
   * into the destination's table. (Within a single workbook the
   * style id can be reused as-is.)
   */
  readonly styleId?: number;
  /**
   * Fully resolved per-cell style snapshot. Populated only when
   * {@link extractClipboardSnapshot} is invoked with a `styleTable`
   * argument — i.e. by the in-app clipboard producer. Cross-format
   * paste handlers (XLSX → DOCX, XLSX → PPTX) use this to project
   * the source font / colour / weight onto the destination's run
   * properties without depending on the source workbook's style
   * table at paste time.
   *
   * Optional and additive: existing readers (paste-range,
   * round-tripped TSV/HTML clipboard) ignore it.
   */
  readonly effectiveStyle?: EffectiveStyle;
}

/**
 * Self-contained clipboard payload. Captures everything we need to
 * round-trip a region of a sheet through the system clipboard:
 *   - origin range so external HTML carries a hint
 *   - dimensions
 *   - row-major matrix of cells (sparse `null` positions allowed)
 *   - merged-region offsets relative to the top-left of the snapshot
 *
 * The shape is intentionally JSON-serialisable so it can survive a
 * round-trip through `text/plain` (TSV form) or be reconstructed
 * from `text/html` (Excel-style table). The fingerprint hop is
 * implemented at the web layer — the snapshot itself stays pure.
 */
export interface XlsxClipboardSnapshot {
  readonly origin: { readonly sheet: string; readonly range: string };
  readonly width: number;
  readonly height: number;
  readonly cells: ReadonlyArray<ReadonlyArray<XlsxClipboardCell | null>>;
  readonly merges: ReadonlyArray<ClipboardMerge>;
}

/** Merge bounds RELATIVE to the snapshot's top-left, 0-based inclusive. */
export interface ClipboardMerge {
  readonly r0: number;
  readonly c0: number;
  readonly r1: number;
  readonly c1: number;
}

/**
 * Project a sheet's range into a `XlsxClipboardSnapshot`.
 *
 * - `range` is an A1 range string (e.g. `"A1:C3"`); single cells
 *   like `"B2"` are also accepted and become a 1×1 snapshot.
 * - Cells outside the populated region come back as `null`.
 * - Merges are clipped to the requested range (Excel parity: copying
 *   half a merge copies the visible half as plain cells).
 */
export function extractClipboardSnapshot(
  sheet: Sheet,
  range: string,
  styleTable?: StyleTable,
): XlsxClipboardSnapshot {
  const r = range.includes(":") ? parseRange(range) : single(range);
  const r0 = Math.min(r.start.row, r.end.row);
  const r1 = Math.max(r.start.row, r.end.row);
  const c0 = Math.min(r.start.col, r.end.col);
  const c1 = Math.max(r.start.col, r.end.col);
  const height = r1 - r0 + 1;
  const width = c1 - c0 + 1;

  const cells: Array<Array<XlsxClipboardCell | null>> = [];
  for (let dr = 0; dr < height; dr++) {
    const row: Array<XlsxClipboardCell | null> = [];
    for (let dc = 0; dc < width; dc++) {
      const cell = sheet.cells.get(cellKey(r0 + dr, c0 + dc));
      if (!cell) {
        row.push(null);
        continue;
      }
      const out: { -readonly [K in keyof XlsxClipboardCell]: XlsxClipboardCell[K] } = {
        value: cell.value,
      };
      if (cell.formula) out.formula = cell.formula.text;
      if (cell.styleId !== undefined) out.styleId = cell.styleId;
      if (styleTable && cell.styleId !== undefined) {
        out.effectiveStyle = flattenCellXf(styleTable, cell.styleId);
      }
      row.push(out);
    }
    cells.push(row);
  }

  const merges: ClipboardMerge[] = [];
  for (const m of sheet.merges) {
    if (m.r2 < r0 || m.r1 > r1 || m.c2 < c0 || m.c1 > c1) continue;
    // Only carry merges that are FULLY contained in the snapshot —
    // a half-clipped merge round-trips as plain cells (matches
    // Excel's own clipboard behaviour and keeps the paste handler
    // from accidentally producing partial-overlap rejections).
    if (m.r1 < r0 || m.r2 > r1 || m.c1 < c0 || m.c2 > c1) continue;
    merges.push({ r0: m.r1 - r0, c0: m.c1 - c0, r1: m.r2 - r0, c1: m.c2 - c0 });
  }

  return {
    origin: {
      sheet: sheet.name,
      range: formatRange({ start: { row: r0, col: c0 }, end: { row: r1, col: c1 } }),
    },
    width,
    height,
    cells,
    merges,
  };
}

function single(ref: string): { start: { row: number; col: number }; end: { row: number; col: number } } {
  const a = parseA1(ref);
  return { start: a, end: a };
}

/**
 * Convert a clipboard snapshot to a TSV (`text/plain`) string.
 * Formula cells render their formula text (with a leading `=`),
 * matching Excel's TSV export. Errors render their `#code`.
 *
 * Cells with embedded tabs / newlines are quoted by escaping —
 * single-line text only for P0; multi-line cells stay as a single
 * cell where the embedded `\n` becomes a literal space (TSV doesn't
 * have a portable quoting convention).
 */
export function snapshotToTsv(snap: XlsxClipboardSnapshot): string {
  const out: string[] = [];
  for (let r = 0; r < snap.height; r++) {
    const row = snap.cells[r] ?? [];
    const cols: string[] = [];
    for (let c = 0; c < snap.width; c++) {
      const cell = row[c];
      cols.push(cell ? formatCellForTsv(cell) : "");
    }
    out.push(cols.join("\t"));
  }
  return out.join("\n");
}

function formatCellForTsv(cell: XlsxClipboardCell): string {
  if (cell.formula) return `=${cell.formula}`;
  return formatValueForTsv(cell.value);
}

function formatValueForTsv(value: CellValue): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return sanitiseTsvString(value);
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "object" && "kind" in value && value.kind === "error") return value.code;
  return String(value);
}

function sanitiseTsvString(s: string): string {
  // Newlines, tabs, and CRs collapse to single spaces — matches Excel's
  // own clipboard TSV behaviour for cells that contain rich text.
  return s.replace(/[\t\r\n]+/g, " ");
}

/**
 * Inverse of `snapshotToTsv`. Splits on rows then on tabs. Cells
 * starting with `=` become formula cells; other cells get their
 * value coerced (number / boolean / string) the same way the
 * formula bar's `parseLiteral` handles literals.
 *
 * Origin and merges are not recoverable from TSV alone — the result
 * sets `origin.range = "A1:..."` and `merges = []`.
 */
export function tsvToSnapshot(text: string): XlsxClipboardSnapshot {
  const rows = text.replace(/\r\n?/g, "\n").split("\n");
  // Excel always appends a trailing newline; drop a single empty
  // trailing row so a single-cell paste doesn't become 1×N rows.
  while (rows.length > 1 && rows[rows.length - 1] === "") rows.pop();
  const split = rows.map((r) => r.split("\t"));
  const height = split.length;
  const width = split.reduce((m, r) => Math.max(m, r.length), 0);
  const cells: Array<Array<XlsxClipboardCell | null>> = [];
  for (let r = 0; r < height; r++) {
    const row: Array<XlsxClipboardCell | null> = [];
    for (let c = 0; c < width; c++) {
      const raw = split[r]?.[c] ?? "";
      row.push(parseClipboardCellLiteral(raw));
    }
    cells.push(row);
  }
  const range = formatRange({
    start: { row: 0, col: 0 },
    end: { row: Math.max(0, height - 1), col: Math.max(0, width - 1) },
  });
  return {
    origin: { sheet: "", range },
    width,
    height,
    cells,
    merges: [],
  };
}

function parseClipboardCellLiteral(raw: string): XlsxClipboardCell | null {
  if (raw === "") return null;
  if (raw.startsWith("=")) return { value: null, formula: raw.slice(1) };
  if (/^-?\d+(?:\.\d+)?$/.test(raw)) {
    const n = Number(raw);
    if (Number.isFinite(n)) return { value: n };
  }
  const lower = raw.toLowerCase();
  if (lower === "true") return { value: true };
  if (lower === "false") return { value: false };
  return { value: raw };
}

/**
 * Pick a delimiter for a `text/plain` payload. Tabs win whenever
 * present (Excel/Sheets always use tabs for cells); otherwise,
 * counts among `,` `;` `|` decide. Falls back to `\t` when the
 * payload is single-cell so the parser still produces a 1×1 grid.
 */
export function sniffDelimiter(text: string): "\t" | "," | ";" | "|" {
  const sample = text.slice(0, 4096);
  const counts = {
    "\t": (sample.match(/\t/g) ?? []).length,
    ",": (sample.match(/,/g) ?? []).length,
    ";": (sample.match(/;/g) ?? []).length,
    "|": (sample.match(/\|/g) ?? []).length,
  };
  if (counts["\t"] > 0) return "\t";
  let best: "\t" | "," | ";" | "|" = "\t";
  let bestN = 0;
  for (const k of [",", ";", "|"] as const) {
    if (counts[k] > bestN) {
      best = k;
      bestN = counts[k];
    }
  }
  return best;
}

/**
 * RFC 4180-ish delimited parser. Supports double-quoted fields with
 * `""` escape, embedded newlines inside quotes, and the four
 * delimiters returned by `sniffDelimiter`. Bare `\n` outside quotes
 * separates rows.
 */
export function delimitedToSnapshot(text: string, delimiter: string): XlsxClipboardSnapshot {
  const rows: string[][] = [[]];
  let field = "";
  let inQuotes = false;
  let i = 0;
  while (i < text.length) {
    const ch = text[i]!;
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"') {
      inQuotes = true;
      i++;
      continue;
    }
    if (ch === delimiter) {
      rows[rows.length - 1]!.push(field);
      field = "";
      i++;
      continue;
    }
    if (ch === "\n" || ch === "\r") {
      rows[rows.length - 1]!.push(field);
      field = "";
      // Swallow CRLF as one separator.
      if (ch === "\r" && text[i + 1] === "\n") i += 2;
      else i++;
      // Don't push a trailing empty row (matches `tsvToSnapshot`).
      if (i < text.length) rows.push([]);
      continue;
    }
    field += ch;
    i++;
  }
  // Push the trailing field ONLY when the file did not end with a
  // row terminator. Otherwise the `\n` already wrote the final
  // field via line 284 and adding `field` here would inject a
  // phantom empty cell at the end of the last row.
  if (field !== "") {
    rows[rows.length - 1]!.push(field);
  }
  // Drop trailing empty row produced by a final newline (kept for
  // safety — `if (i < text.length)` above already prevents most of
  // these, but a CRLF + EOF can leave one behind).
  while (rows.length > 1 && rows[rows.length - 1]!.length === 1 && rows[rows.length - 1]![0] === "") {
    rows.pop();
  }
  while (rows.length > 1 && rows[rows.length - 1]!.length === 0) {
    rows.pop();
  }

  const height = rows.length;
  const width = rows.reduce((m, r) => Math.max(m, r.length), 0);
  const cells: Array<Array<XlsxClipboardCell | null>> = [];
  for (let r = 0; r < height; r++) {
    const row: Array<XlsxClipboardCell | null> = [];
    for (let c = 0; c < width; c++) {
      const raw = rows[r]?.[c] ?? "";
      row.push(parseClipboardCellLiteral(raw));
    }
    cells.push(row);
  }
  const range = formatRange({
    start: { row: 0, col: 0 },
    end: { row: Math.max(0, height - 1), col: Math.max(0, width - 1) },
  });
  return {
    origin: { sheet: "", range },
    width,
    height,
    cells,
    merges: [],
  };
}
