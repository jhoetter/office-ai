import type { XlsxClipboardSnapshot } from "./snapshot.js";
import { delimitedToSnapshot, sniffDelimiter, tsvToSnapshot } from "./snapshot.js";

/**
 * Headless ("no DOM") parsers for clipboard payloads originating
 * outside our editor — Excel-on-web, Excel-on-desktop, Google Sheets,
 * Numbers, plain text editors, etc.
 *
 * Why a hand-rolled parser instead of `DOMParser`?
 *   - The `@officeai/xlsx` package is intentionally DOM-free
 *     (`Headless-first Design` pillar) — it must run in Node CI,
 *     workers, and the agent CLI. Bringing in `jsdom` would break
 *     that.
 *   - The actual subset we need is tiny: `<table>`, `<tr>`, `<td>`
 *     / `<th>`, plus the `data-xlsx-fingerprint` round-trip
 *     attribute. A 200-line regex/state-machine parser handles every
 *     fixture in `__fixtures__/`.
 *
 * Outbound marshalling lives in `apps/web/app/xlsx-editor/clipboard.ts`
 * because it talks to the browser's `ClipboardItem` API. That file
 * delegates *parsing* back here so Excel-fixture coverage is in the
 * package test suite where it belongs.
 */

const FINGERPRINT_RE = /<table[^>]*\sdata-xlsx-fingerprint="([^"]+)"/i;

/**
 * Top-level clipboard parser. Tries, in priority order:
 *   1. Our own HTML fingerprint (full fidelity round-trip).
 *   2. Generic HTML `<table>` (Excel, Sheets, Numbers, Word).
 *   3. Plain text (TSV, CSV, semicolon, pipe).
 *
 * Returns `null` when no payload variant could be parsed.
 */
export function parseExternalClipboard(opts: {
  readonly html?: string | null;
  readonly text?: string | null;
}): XlsxClipboardSnapshot | null {
  if (opts.html) {
    const fp = parseFingerprintHtml(opts.html);
    if (fp) return fp;
    const tbl = parseHtmlTable(opts.html);
    if (tbl) return tbl;
  }
  if (opts.text) {
    const delim = sniffDelimiter(opts.text);
    return delim === "\t" ? tsvToSnapshot(opts.text) : delimitedToSnapshot(opts.text, delim);
  }
  return null;
}

/**
 * Recover an exact {@link XlsxClipboardSnapshot} from the
 * `data-xlsx-fingerprint` attribute on a `<table>` element. This is
 * how a copy-then-paste inside our own app preserves formulas,
 * styleIds, and merge regions losslessly.
 */
export function parseFingerprintHtml(html: string): XlsxClipboardSnapshot | null {
  const m = FINGERPRINT_RE.exec(html);
  if (!m) return null;
  try {
    const decoded = decodeURIComponent(m[1]!);
    const snap = JSON.parse(decoded) as unknown;
    if (
      snap &&
      typeof snap === "object" &&
      "cells" in snap &&
      Array.isArray((snap as { cells: unknown[] }).cells) &&
      "width" in snap &&
      "height" in snap
    ) {
      return snap as XlsxClipboardSnapshot;
    }
  } catch {
    // fall through
  }
  return null;
}

/**
 * Headless `<table>` parser. Walks the document with a small state
 * machine that tracks the current open tag and its text buffer.
 * Handles:
 *   - `<table>` / `<tbody>` / `<thead>` / `<tfoot>` containers
 *   - `<tr>` rows
 *   - `<td>` / `<th>` cells (including `colspan` attribute)
 *   - HTML entities (numeric `&#nn;` and the common named ones)
 *   - Self-closing tags (`<br/>` becomes a newline inside a cell)
 *
 * Cells inside the first `<table>` only — Office HTML often nests
 * outer wrapper tables for styling, but the *data* table is the one
 * with the most rows; we pick the largest by cell-count.
 */
export function parseHtmlTable(html: string): XlsxClipboardSnapshot | null {
  if (!html) return null;
  const tables = extractAllTables(html);
  if (tables.length === 0) return null;
  let best = tables[0]!;
  for (const t of tables) {
    if (t.cells.length * (t.cells[0]?.length ?? 0) > best.cells.length * (best.cells[0]?.length ?? 0)) {
      best = t;
    }
  }
  const width = best.cells.reduce((m, r) => Math.max(m, r.length), 0);
  const cells = best.cells.map((row) => {
    const out: Array<{ value: number | string | boolean | null } | null> = [];
    for (let c = 0; c < width; c++) {
      const raw = (row[c] ?? "").trim();
      if (raw === "") {
        out.push(null);
        continue;
      }
      if (/^-?\d+(?:\.\d+)?$/.test(raw)) {
        const n = Number(raw);
        if (Number.isFinite(n)) {
          out.push({ value: n });
          continue;
        }
      }
      const lower = raw.toLowerCase();
      if (lower === "true") {
        out.push({ value: true });
        continue;
      }
      if (lower === "false") {
        out.push({ value: false });
        continue;
      }
      out.push({ value: raw });
    }
    return out;
  });
  return {
    origin: { sheet: "", range: "" },
    width,
    height: cells.length,
    cells,
    merges: [],
  };
}

interface ParsedTable {
  readonly cells: ReadonlyArray<ReadonlyArray<string>>;
}

function extractAllTables(html: string): ParsedTable[] {
  const out: ParsedTable[] = [];
  const tableRe = /<table\b[^>]*>([\s\S]*?)<\/table>/gi;
  let tm: RegExpExecArray | null;
  while ((tm = tableRe.exec(html)) !== null) {
    out.push({ cells: parseTableBody(tm[1]!) });
  }
  return out;
}

function parseTableBody(body: string): string[][] {
  const rows: string[][] = [];
  const rowRe = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi;
  let rm: RegExpExecArray | null;
  while ((rm = rowRe.exec(body)) !== null) {
    rows.push(parseRow(rm[1]!));
  }
  return rows;
}

function parseRow(rowHtml: string): string[] {
  const out: string[] = [];
  // Match <td> or <th> with optional attributes; capture inner text.
  const cellRe = /<(td|th)\b([^>]*)>([\s\S]*?)<\/\1>/gi;
  let cm: RegExpExecArray | null;
  while ((cm = cellRe.exec(rowHtml)) !== null) {
    const attrs = cm[2] ?? "";
    const inner = cm[3] ?? "";
    const text = stripTagsAndDecode(inner);
    out.push(text);
    // Honour `colspan="N"` by emitting N-1 empty trailing cells so
    // the column index of the next real cell stays right.
    const csm = /\bcolspan\s*=\s*["']?(\d+)/i.exec(attrs);
    if (csm) {
      const n = Math.max(1, Math.min(64, Number(csm[1])));
      for (let i = 1; i < n; i++) out.push("");
    }
  }
  return out;
}

function stripTagsAndDecode(html: string): string {
  // Replace <br>, <br/>, <br /> with newline before stripping tags so
  // multi-line cells round-trip.
  let s = html.replace(/<br\s*\/?\s*>/gi, "\n");
  // Office sometimes wraps cell content in a <p> or <span>; both can
  // safely be treated as their inner text.
  s = s.replace(/<\/?(?:p|span|div|font|b|i|u|strong|em)\b[^>]*>/gi, "");
  // Strip remaining tags conservatively.
  s = s.replace(/<[^>]+>/g, "");
  s = decodeEntities(s);
  return s;
}

function decodeEntities(s: string): string {
  return s
    .replace(/&#(\d+);/g, (_m, n: string) => String.fromCharCode(Number(n)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_m, n: string) => String.fromCharCode(parseInt(n, 16)))
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;|&apos;/g, "'");
}
